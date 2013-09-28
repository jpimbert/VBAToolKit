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
' Purpose   : Test the vtkExportConfiguration function and the BeforeSave handler using this function.
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

Private Const testProjectName As String = "TestProject"
Private Const testProjectFileName As String = testProjectName & "_DEV.xlsm"
Private Const testModule1FileName As String = "TestModule1.bas"
Private Const testModule2FileName As String = "TestModule2.bas"
Private Const importedModuleFileName As String = "TestModule.bas"
Private Const testModule1Name As String = "TestModule1"
Private Const testModule2Name As String = "TestModule2"
Private Const importedModuleName As String = "TestModule"
Private Const totalNumberOfModules As Integer = 2

Private Const newlyCreatedProjectName As String = "NewlyCreated"
Private Const newlyCreatedProjectFileName As String = newlyCreatedProjectName & "_DEV.xlsm"

Private cm As vtkConfigurationManager
Private testConf As vtkConfiguration

Private testWorkbook As Workbook
Private newlyCreatedWorkbook As Workbook

Private fso As New FileSystemObject

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
    
    Set testWorkbook = VBAToolKit.getTestFileFromTemplate(fileName:=testProjectFileName, openExcel:=True)
    
    Set cm = vtkConfigurationManagerForProject(testProjectName)
    Set testConf = cm.configurations(testWorkbook.VBProject.name)
    
End Sub

Private Sub ITestCase_TearDown()
    On Error Resume Next
    
    vtkResetConfigurationManagers
    vtkCloseAndKillWorkbook wb:=testWorkbook
    vtkCloseAndKillWorkbook wb:=newlyCreatedWorkbook
    vtkCleanFolder VBAToolKit.vtkTestPath

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
'    - expected : 0 module is exported since it has not been added to the configuration
'    - Test_vtkExportConfiguration_importedModuleMustNotBeExported
'
' - Source files are exported, 1 module is imported, added to the configuration manager, onlyModified = True
'    - expected : 0 module is exported since it has not been added to the configuration itself.
'    - Test_vtkExportConfiguration_importedModuleAddeToConfigurationManagerMustNotBeExported
'
' - Source files are exported, 1 module is modified, onlyModified = True
'    - expected : one module is exported and modified
'
' - Souce files are exported, 1 module is modified, onlyModified = True
'    - expected : 1 module is exported, the unexported modules are not normalized
'
' - Source files are exported, no module is modified, onlyModified = False
'    - expected : all modules are exported, already existing modules are normalized
'
' -------------------------------------------------------------------------


Public Sub Test_vtkExportConfiguration_filesNotYetCreatedTrue()
' Source files are not yet created, onlyModified = True => expected : all modules are exported
    
    Dim retval As Integer
    
    On Error GoTo Test_vtkExportConfiguration_filesNotYetCreatedTrue_Error

    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=True)
    mAssert.Equals retval, totalNumberOfModules, totalNumberOfModules & " should be exported as files do not yet exist."
    
    mAssert.Should fso.fileExists(VBAToolKit.vtkTestPath & "\" & testModule1FileName), testModule1FileName & " should have been created."
    mAssert.Should fso.fileExists(VBAToolKit.vtkTestPath & "\" & testModule2FileName), testModule2FileName & " should have been created."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesNotYetCreatedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesNotYetCreatedTrue of Class Module vtkExportConfigurationTester"
    Resume Next
    
End Sub

Public Sub Test_vtkExportConfiguration_filesNotYetCreatedFalse()
' Source files are not yet created, onlyModified = False => expected : all modules are exported
    
    Dim retval As Integer

    On Error GoTo Test_vtkExportConfiguration_filesNotYetCreatedFalse_Error

    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=False)
    mAssert.Equals retval, totalNumberOfModules, totalNumberOfModules & " should be exported as files do not yet exist."

    mAssert.Should fso.fileExists(VBAToolKit.vtkTestPath & "\" & testModule1FileName), testModule1FileName & " should have been created."
    mAssert.Should fso.fileExists(VBAToolKit.vtkTestPath & "\" & testModule2FileName), testModule2FileName & " should have been created."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesNotYetCreatedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesNotYetCreatedFalse of Class Module vtkExportConfigurationTester"
    Resume Next

End Sub


Public Sub Test_vtkExportConfiguration_filesJustExportedTrue()
' Source files are created, onlyModified = True => expected : 0 modules are exported
    
    Dim retval As Integer
    Dim oldFileSize1 As Long
    Dim newFileSize1 As Long
    Dim oldFileSize2 As Long
    Dim newFileSize2 As Long
    
    On Error GoTo Test_vtkExportConfiguration_filesYetCreatedTrue_Error

    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName

    oldFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    oldFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size

    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=True)
    
    newFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    newFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size
    
    mAssert.Equals retval, 0, "No module should be exported as files yet exist and onlyModified = True"

    mAssert.Should oldFileSize1 = newFileSize1, testModule1FileName & " must be the same size as before export as the corresponding module has not been exported."
    mAssert.Should oldFileSize2 = newFileSize2, testModule2FileName & " must be the same size as before export as the corresponding module has not been exported."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesYetCreatedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesYetCreatedTrue of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub

Public Sub Test_vtkExportConfiguration_filesJustExportedFalse()
' Source files are created, onlyModified = False => expected : all modules are exported
    
    Dim retval As Integer
    Dim oldFileSize1 As Long
    Dim newFileSize1 As Long
    Dim oldFileSize2 As Long
    Dim newFileSize2 As Long
    
    On Error GoTo Test_vtkExportConfiguration_filesJustExportedFalse_Error
    
    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName

    oldFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    oldFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size

    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=False)
    
    newFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    newFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size
    
    mAssert.Equals retval, totalNumberOfModules, totalNumberOfModules & " modules should be exported as onlyModified = False."

    mAssert.Should oldFileSize1 = newFileSize1, testModule1FileName & " must be the same size after exporting since the corresponding module has not been modified."
    mAssert.Should oldFileSize2 = newFileSize2, testModule2FileName & " must be the same size after exporting since the corresponding module has not been modified."


    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesJustExportedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesJustExportedFalse of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub

Public Sub Test_vtkExportConfiguration_oneFileModifiedTrue()
' Source files are exported, 1 module is modified, onlyModified = True => expected : 1 module is exported

    Dim retval As Integer
    Dim oldFileSize1 As Long
    Dim newFileSize1 As Long
    Dim oldFileSize2 As Long
    Dim newFileSize2 As Long
    
    On Error GoTo Test_vtkExportConfiguration_oneFileModifiedTrue_Error

    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName

    oldFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    oldFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size

    insertDummyProcedureInCodeModule testWorkbook.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=1
    
    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=True)
    
    newFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    newFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size
    
    mAssert.Equals retval, 1, "1 module should be exported as 1 module has been modified and onlyModified = True"
    mAssert.Should oldFileSize1 < newFileSize1, testModule1FileName & " must be bigger after adding code in the module and exporting."
    mAssert.Should oldFileSize2 = newFileSize2, testModule2FileName & " must be the same after exporting since it has not been modified."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_oneFileModifiedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_oneFileModifiedTrue of Class Module vtkExportConfigurationTester"
    Resume Next
    
End Sub


Public Sub Test_vtkExportConfiguration_oneFileModifiedFalse()
' Source files are exported, 1 module is modified, onlyModified = False => expected : all modules are exported
    
    Dim retval As Integer
    Dim oldFileSize1 As Long
    Dim newFileSize1 As Long
    Dim oldFileSize2 As Long
    Dim newFileSize2 As Long
    
    On Error GoTo Test_vtkExportConfiguration_oneFileModifiedFalse_Error

    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName

    oldFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    oldFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size

    insertDummyProcedureInCodeModule testWorkbook.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=1
    
    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=False)
    
    newFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    newFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size
    
    mAssert.Equals retval, totalNumberOfModules, totalNumberOfModules & " modules should be exported as onlyModified = False"
    
    mAssert.Should oldFileSize1 < newFileSize1, testModule1FileName & " must be bigger after adding code in the module and exporting."
    mAssert.Should oldFileSize2 = newFileSize2, testModule2FileName & " must be the same after exporting since it has not been modified."
    
    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_oneFileModifiedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_oneFileModifiedFalse of Class Module vtkExportConfigurationTester"
    Resume Next
    
End Sub


Public Sub Test_vtkExportConfiguration_twoFilesModifiedTrue()
' Source files are exported, 2 modules are modified, onlyModified = True => expected : 2 modules are exported
    
    Dim retval As Integer
    Dim oldFileSize1 As Long
    Dim newFileSize1 As Long
    Dim oldFileSize2 As Long
    Dim newFileSize2 As Long
    
    On Error GoTo Test_vtkExportConfiguration_twoFilesModifiedTrue_Error

    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName
    
    oldFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    oldFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size
    
    insertDummyProcedureInCodeModule testWorkbook.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=1
    insertDummyProcedureInCodeModule testWorkbook.VBProject.VBComponents(testModule2Name).CodeModule, dummyInt:=2
    
    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=True)
    
    newFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    newFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size
    
    mAssert.Equals retval, 2, "2 modules should be exported as 2 modules have been modified and onlyModified = True"

    mAssert.Should oldFileSize1 < newFileSize1, testModule1FileName & " must be bigger after adding code in the module and exporting."
    mAssert.Should oldFileSize2 < newFileSize2, testModule2FileName & " must be bigger after adding code in the module and exporting."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_twoFilesModifiedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_twoFilesModifiedTrue of Class Module vtkExportConfigurationTester"
    Resume Next

End Sub

Public Sub Test_vtkExportConfiguration_twoFilesModifiedFalse()
' Source files are exported, 2 modules are modified, onlyModified = False => expected : all modules are exported
    
    Dim retval As Integer
    Dim oldFileSize1 As Long
    Dim newFileSize1 As Long
    Dim oldFileSize2 As Long
    Dim newFileSize2 As Long
    
    On Error GoTo Test_vtkExportConfiguration_twoFilesModifiedFalse_Error

    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName

    oldFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    oldFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size

    insertDummyProcedureInCodeModule testWorkbook.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=1
    insertDummyProcedureInCodeModule testWorkbook.VBProject.VBComponents(testModule2Name).CodeModule, dummyInt:=2
    
    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=False)
    
    newFileSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    newFileSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size
    
    mAssert.Equals retval, totalNumberOfModules, totalNumberOfModules & " modules should be exported as onlyModified = False"

    mAssert.Should oldFileSize1 < newFileSize1, testModule1FileName & " must be bigger after adding code in the module and exporting."
    mAssert.Should oldFileSize2 < newFileSize2, testModule2FileName & " must be bigger after adding code in the module and exporting."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_twoFilesModifiedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_twoFilesModifiedFalse of Class Module vtkExportConfigurationTester"
    Resume Next

End Sub

Public Sub Test_vtkExportConfiguration_notReferencedModuleMustNotBeExported()
' Source files are exported, 1 module is imported, onlyModified = True => expected : no module is exported
' since no module has been added to the configuration.
    
    Dim retval As Integer

    On Error GoTo Test_vtkExportConfiguration_notReferencedModuleMustNotBeExported_Error

    ' Export the modules already present so that they don't get exported again
    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName

    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=True)

    ' We import the module
    vtkImportOneModule testWorkbook.VBProject, importedModuleName, VBAToolKit.vtkTestPath & "\" & importedModuleFileName
        
    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=True)
    
    mAssert.Equals retval, 0, "The module should not be exported as it has not been added to the configuration."
    
    mAssert.Should Not fso.fileExists(VBAToolKit.vtkTestPath & "\" & importedModuleFileName), importedModuleName & " should not have been created."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_notReferencedModuleMustNotBeExported_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_importedFileMustBeExported of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub

Public Sub Test_vtkExportConfiguration_importedModuleAddedToConfigurationManagerMustNotBeExported()
' Source files are exported, 1 module is imported, added to the configuration manager, onlyModified = True => expected :
' 0 module is exported since it has not been added to the configuration itself.
    
    Dim retval As Integer

    On Error GoTo Test_vtkExportConfiguration_importedModuleAddedToConfigurationManagerMustBeExported_Error

    ' Export the modules already present so that they don't get exported again
    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName

    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=True)

    ' We import the module
    vtkImportOneModule testWorkbook.VBProject, importedModuleName, VBAToolKit.vtkTestPath & "\" & importedModuleFileName
    cm.addModule (importedModuleName)
    
    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=True)
    
    mAssert.Equals retval, 0, "No module should be exported as none has been added to the configuration."
    
    mAssert.Should Not fso.fileExists(VBAToolKit.vtkTestPath & "\" & importedModuleFileName), importedModuleName & " should not have been created."

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_importedModuleAddedToConfigurationManagerMustBeExported_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_importedFileMustBeExported of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub


Public Sub Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook()
' adding the handler in a valid workbook, modifying one module, saving => 1 module is exported
    
    Dim retval As Integer
    Dim oldSize1 As Integer
    Dim oldSize2 As Integer
    Dim newSize1 As Integer
    Dim newSize2 As Integer

    On Error GoTo Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook_Error
              
    ' Export the modules already in the project
    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName

    ' Get the size of the exported modules before the save
    oldSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    oldSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size

    ' Add the handler
    vtkAddBeforeSaveHandlerInDEVWorkbook testWorkbook, testProjectName, testConf.name

    ' Modify a module
    insertDummyProcedureInCodeModule testWorkbook.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=1

    ' Save the workbook
    testWorkbook.Save

    ' Get the new size of the exported modules
    newSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    newSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size

    mAssert.Should oldSize1 < newSize1, testModule1FileName & " must be bigger after the save."
    mAssert.Should oldSize2 = newSize2, testModule2FileName & " must be of equal size after the save."

    On Error GoTo 0
    Exit Sub

Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook of Class Module vtkBeforeSaveHandlerTester"
    ITestCase_TearDown
    Exit Sub
    
End Sub

'-----------------------------------------------
' TESTS REGARDING NORMALIZATION
'-----------------------------------------------

Private Sub insertDummyLineInCodeModule(codemo As VBIDE.CodeModule, dummyLine As String)
    With codemo
        .InsertLines .CountOfLines + 1, dummyLine
    End With
End Sub

Private Function getLastLineOfTextFile(textFilePath As String) As String
    Dim str As String
    Dim tmp As String
    
    str = vtkTextFileReader(textFilePath)
    tmp = Split(str, Chr(10))(UBound(Split(str, Chr(10))))
    If Len(tmp) < 2 Then tmp = Split(str, Chr(10))(UBound(Split(str, Chr(10))) - 1)
    getLastLineOfTextFile = tmp
    
End Function


Public Sub Test_vtkExportConfiguration_modifiedModuleIsNormalizedOMTrue()
' - Source files are exported, 1 module is modified, onlyModified = True
'    - expected : one module is exported and modified

    Dim retval As Integer

    On Error GoTo vtkExportConfiguration_modifiedModuleIsNormalizedOMTrue_Error

    ' Export the modules already present so that they don't get exported again
    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName
    
    ' Insert a line that contains key words badly cased
    insertDummyLineInCodeModule testWorkbook.VBProject.VBComponents(testModule1Name).CodeModule, "err dim"
    
    ' Export the modified modules
    retval = vtkExportConfiguration(testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=True)
    
    ' Check the casing of the last line of the modified module
    Dim expected As String
    Dim result As String
    expected = "Err Dim"
    result = getLastLineOfTextFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName)
    
    mAssert.Should StrComp(expected, result) = 0, _
                    "File should have been formatted." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result   : " & result
    
    ' Last check
    mAssert.Equals retval, 1, "Only one module should have been exported."
    
    On Error GoTo 0
    Exit Sub

vtkExportConfiguration_modifiedModuleIsNormalizedOMTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure vtkExportConfiguration_modifiedModuleIsNormalizedTure of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub

Public Sub Test_vtkExportConfiguration_notModifiedModulesAreNotNormalizedOMTrue()
' - Souce files are exported, 1 module is modified, onlyModified = True
'    - expected : 1 module is exported, the unexported modules are not normalized

    On Error GoTo vtkExportConfiguration_notModifiedModulesAreNotNormalizedOMTrue_Error

    ' Export the modules already present so that they don't get exported again
    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName, normalize:=False
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName, normalize:=False

    ' Insert a line in code module 1 that contains key words badly cased
    insertDummyLineInCodeModule testWorkbook.VBProject.VBComponents(testModule1Name).CodeModule, "dim retval"

    ' Export the modified modules
    vtkExportConfiguration testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=True

    ' Check the casing of the last line of the untouched module. This line contained "Dim retval" before the export.
    Dim expected As String
    Dim result As String
    expected = "Dim retval"
    result = getLastLineOfTextFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName)

    mAssert.Should StrComp(expected, result) = 0, _
                    "File should not have been fomatted as it has not been modified." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result   : " & result

    On Error GoTo 0
    Exit Sub
vtkExportConfiguration_notModifiedModulesAreNotNormalizedOMTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure vtkExportConfiguration_notModifiedModulesAreNotNormalizedOMTrue of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub


Public Sub Test_vtkExportConfiguration_notModifiedModulesAreNormalizedOMFalse()
' - Source files are exported, no module is modified, onlyModified = False
'    - expected : all modules are exported, already existing modules are normalized

    On Error GoTo vtkExportConfiguration_notModifiedModulesAreNormalizedOMFalse

    ' Export the modules already present so that they don't get exported again
    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName

    ' Insert a line in code module 1 that contains key words badly cased
    insertDummyLineInCodeModule testWorkbook.VBProject.VBComponents(testModule1Name).CodeModule, "err dim"

    ' Export all the modified modules
    vtkExportConfiguration testWorkbook.VBProject, testProjectName, testConf.name, onlyModified:=False

    ' Check the casing of the last line of the untouched module. This line contained "Dim retval" before the export.
    Dim expected As String
    Dim result As String
    expected = "Dim retVal"
    result = getLastLineOfTextFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName)

    mAssert.Should StrComp(expected, result) = 0, _
                    "File should have been formatted as all modules were formatted." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result   : " & result

    ' Check the casing of the last line of the modified module.
    expected = "Err Dim"
    result = getLastLineOfTextFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName)

    mAssert.Should StrComp(expected, result) = 0, _
                    "File should not have been formatted as all modules were formatted." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result   : " & result

    On Error GoTo 0
    Exit Sub

vtkExportConfiguration_notModifiedModulesAreNormalizedOMFalse:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure vtkExportCOnfiguration_notModifiedModulesAreNormalizedFalse of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_filesNotYetCreatedTrue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_filesNotYetCreatedFalse"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_filesJustExportedTrue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_filesJustExportedFalse"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_oneFileModifiedTrue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_oneFileModifiedFalse"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_twoFilesModifiedTrue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_twoFilesModifiedFalse"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_notReferencedModuleMustNotBeExported"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_importedModuleAddedToConfigurationManagerMustNotBeExported"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_modifiedModuleIsNormalizedOMTrue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_notModifiedModulesAreNotNormalizedOMTrue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfiguration_notModifiedModulesAreNormalizedOMFalse"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkExportConfiguration_filesNotYetCreatedTrue": Test_vtkExportConfiguration_filesNotYetCreatedTrue
        Case "Test_vtkExportConfiguration_filesNotYetCreatedFalse": Test_vtkExportConfiguration_filesNotYetCreatedFalse
        Case "Test_vtkExportConfiguration_filesJustExportedTrue": Test_vtkExportConfiguration_filesJustExportedTrue
        Case "Test_vtkExportConfiguration_filesJustExportedFalse": Test_vtkExportConfiguration_filesJustExportedFalse
        Case "Test_vtkExportConfiguration_oneFileModifiedTrue": Test_vtkExportConfiguration_oneFileModifiedTrue
        Case "Test_vtkExportConfiguration_oneFileModifiedFalse": Test_vtkExportConfiguration_oneFileModifiedFalse
        Case "Test_vtkExportConfiguration_twoFilesModifiedTrue": Test_vtkExportConfiguration_twoFilesModifiedTrue
        Case "Test_vtkExportConfiguration_twoFilesModifiedFalse": Test_vtkExportConfiguration_twoFilesModifiedFalse
        Case "Test_vtkExportConfiguration_notReferencedModuleMustNotBeExported": Test_vtkExportConfiguration_notReferencedModuleMustNotBeExported
        Case "Test_vtkExportConfiguration_importedModuleAddedToConfigurationManagerMustNotBeExported": Test_vtkExportConfiguration_importedModuleAddedToConfigurationManagerMustNotBeExported
        Case "Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook": Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook
        Case "Test_vtkExportConfiguration_modifiedModuleIsNormalizedOMTrue": Test_vtkExportConfiguration_modifiedModuleIsNormalizedOMTrue
        Case "Test_vtkExportConfiguration_notModifiedModulesAreNotNormalizedOMTrue": Test_vtkExportConfiguration_notModifiedModulesAreNotNormalizedOMTrue
        Case "Test_vtkExportConfiguration_notModifiedModulesAreNormalizedOMFalse": Test_vtkExportConfiguration_notModifiedModulesAreNormalizedOMFalse
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


