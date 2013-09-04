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
Private Const testModule1 As String = "TestModule1.bas"
Private Const testModule2 As String = "TestModule2.bas"
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
    On Error GoTo ITestCase_SetUp_Error

    Set mAssert = Assert
    
    FileCopy source:=vtkPathToTemplateFolder & "\" & testWorkbookName, destination:=vtkPathToTestFolder & "\" & testWorkbookName

    Set wbForTest = Workbooks.Open(fileName:=vtkPathToTestFolder & "\" & testWorkbookName)
    
    vtkResetConfigurationManagers
    Set cm = vtkConfigurationManagerForProject(testProjectName)
    Set testConf = cm.configurations(wbForTest.VBProject.name)

    On Error GoTo 0
    Exit Sub

ITestCase_SetUp_Error:
    Debug.Print "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure ITestCase_SetUp of Class Module vtkExportConfigurationTester"
    vtkCloseAndKillWorkbook Wb:=wbForTest
    Exit Sub
    
End Sub

Private Sub ITestCase_TearDown()
    
    vtkCloseAndKillWorkbook Wb:=wbForTest
    vtkCleanFolder vtkPathToTestFolder

End Sub
'--------------------------------------------------
' Tests to be written
'   - export configuration with files not yet created and onlyModified = True => retval = number of modules in the configuration
'   - export configuration with files not yet created and onlyModified = False => retval = number of modules in the configuration
'   - export configuration with files just exported and onlyModified = True => retval = number of modules in the configuration
'   - export configuration with files just exported and onlyModified = False => retval = 0
'
'   - export configuration with one file modified and onlyModified = True => retval = 1
'   - export configuration with one file modified and onlyModified = False => retval = 2
'   - export configuration with two files modified and onlyModified = True => retval = 2
'   - export configuration with two files modified and onlyModified = False => retval = 2
' -------------------------------------------------

Private Sub insertDummyProcedureInCodeModule(codemo As VBIDE.CodeModule, dummyInt As Integer)
    Dim dummyProcedure As String
    
    dummyProcedure = _
    "Public Sub dummyProcedure" & dummyInt & "()" & vbNewLine & _
    "End Sub" & vbNewLine
    
    With codemo
        .InsertLines .CountOfLines + 1, dummyProcedure
    End With

End Sub


Public Sub Test_vtkExportConfiguration_filesNotYetCreatedTrue()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_filesNotYetCreatedTrue_Error

    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=True)
    mAssert.Equals retVal, totalNumberOfModules

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
    mAssert.Equals retVal, totalNumberOfModules

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesNotYetCreatedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesNotYetCreatedFalse of Class Module vtkExportConfigurationTester"
    Resume Next

End Sub

Public Sub Test_vtkExportConfiguration_filesJustExportedTrue()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_filesYetCreatedTrue_Error

    vtkExportOneModule wbForTest.VBProject, testModule1, vtkPathToTestFolder & "\" & testModule1
    vtkExportOneModule wbForTest.VBProject, testModule2, vtkPathToTestFolder & "\" & testModule2

    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=True)
    mAssert.Equals retVal, totalNumberOfModules

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesYetCreatedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesYetCreatedTrue of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub

Public Sub Test_vtkExportConfiguration_filesJustExportedFalse()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_filesJustExportedFalse_Error
    
    vtkExportOneModule wbForTest.VBProject, testModule1, vtkPathToTestFolder & "\" & testModule1
    vtkExportOneModule wbForTest.VBProject, testModule2, vtkPathToTestFolder & "\" & testModule2

    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=False)
    mAssert.Equals retVal, 0

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_filesJustExportedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_filesJustExportedFalse of Class Module vtkExportConfigurationTester"
    Resume Next
End Sub

Public Sub Test_vtkExportConfiguration_oneFileModifiedTrue()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_oneFileModifiedTrue_Error

    vtkExportOneModule wbForTest.VBProject, testModule1, vtkPathToTestFolder & "\" & testModule1
    vtkExportOneModule wbForTest.VBProject, testModule2, vtkPathToTestFolder & "\" & testModule2

    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1).CodeModule, dummyInt:=1
    
    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=True)
    mAssert.Equals retVal, 1

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_oneFileModifiedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_oneFileModifiedTrue of Class Module vtkExportConfigurationTester"
    Resume Next
    
End Sub


Public Sub Test_vtkExportConfiguration_oneFileModifiedFalse()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_oneFileModifiedFalse_Error

    vtkExportOneModule wbForTest.VBProject, testModule1, vtkPathToTestFolder & "\" & testModule1
    vtkExportOneModule wbForTest.VBProject, testModule2, vtkPathToTestFolder & "\" & testModule2

    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1).CodeModule, dummyInt:=1
    
    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=False)
    mAssert.Equals retVal, totalNumberOfModules

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_oneFileModifiedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_oneFileModifiedFalse of Class Module vtkExportConfigurationTester"
    Resume Next
    
End Sub


Public Sub Test_vtkExportConfiguration_twoFilesModifiedTrue()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_twoFilesModifiedTrue_Error

    vtkExportOneModule wbForTest.VBProject, testModule1, vtkPathToTestFolder & "\" & testModule1
    vtkExportOneModule wbForTest.VBProject, testModule2, vtkPathToTestFolder & "\" & testModule2

    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1).CodeModule, dummyInt:=1
    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1).CodeModule, dummyInt:=2
    
    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=True)
    mAssert.Equals retVal, 2

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_twoFilesModifiedTrue_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_twoFilesModifiedTrue of Class Module vtkExportConfigurationTester"
    Resume Next

End Sub

Public Sub Test_vtkExportConfiguration_twoFilesModifiedFalse()
    Dim retVal As Integer
    
    On Error GoTo Test_vtkExportConfiguration_twoFilesModifiedFalse_Error

    vtkExportOneModule wbForTest.VBProject, testModule1, vtkPathToTestFolder & "\" & testModule1
    vtkExportOneModule wbForTest.VBProject, testModule2, vtkPathToTestFolder & "\" & testModule2

    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1).CodeModule, dummyInt:=1
    insertDummyProcedureInCodeModule wbForTest.VBProject.VBComponents(testModule1).CodeModule, dummyInt:=2
    
    retVal = vtkExportConfiguration(wbForTest.VBProject, testProjectName, testConf.name, onlyModified:=False)
    mAssert.Equals retVal, totalNumberOfModules

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfiguration_twoFilesModifiedFalse_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportConfiguration_twoFilesModifiedFalse of Class Module vtkExportConfigurationTester"
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
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

