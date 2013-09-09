VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkBeforeSaveHandlerTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkBeforeSaveHandlerTester
' Author    : Lucas Vitorino
' Purpose   : Test the Workbook_BeforeSave handler :
'               - its addition with vtkAddBeforeSaveHandlerInDEVWorkbook
'               - its behaviour
'
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
Private Const testWorkbookName As String = testProjectName & "_DEV.xlsm"
Private testWorkbook As Workbook

Private Const existingProjectName As String = "ExistingProject"
Private Const existingProjectWorkbookName As String = "WorkBookForCOnfigurationsTests.xlsm"
Private existingProjectWorkbook As Workbook


Private Const testModule1FileName As String = "TestModule1.bas"
Private Const testModule2FileName As String = "TestModule2.bas"
Private Const importedModuleFileName As String = "TestModule.bas"
Private Const testModule1Name As String = "TestModule1"
Private Const testModule2Name As String = "TestModule2"
Private Const importedModuleName As String = "TestModule"
Private Const totalNumberOfModules As Integer = 2

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
    
End Sub

Private Sub ITestCase_TearDown()
    
    On Error Resume Next
    
    vtkCloseAndKillWorkbook wb:=testWorkbook
    vtkCloseAndKillWorkbook wb:=existingProjectWorkbook
    vtkCleanFolder VBAToolKit.vtkTestPath
    
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

'----------------------------------------------------------------
' TESTS to be written
'----------------------------------------------------------------
'
' Adding the handler
'
'   - adding the handler in a valid workbook, modifying one module, saving => 1 module is exported
'   - adding the handler in a workbook whose name is not valid because it is not finished by DEV => forbidden parameter error
'
'----------------------------------------------------------------

Public Sub Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook()
' adding the handler in a valid workbook, modifying one module, saving => 1 module is exported
    
    Dim retVal As Integer
    Dim oldSize1 As Integer
    Dim oldSize2 As Integer
    Dim newSize1 As Integer
    Dim newSize2 As Integer
    
    On Error GoTo Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook_Error
    
    ' Open the relevant workbook from the templates
    Set testWorkbook = VBAToolKit.getTestFileFromTemplate(fileName:=testWorkbookName, openExcel:=True)
    
    ' Export the modules already in the project
    vtkExportOneModule testWorkbook.VBProject, testModule1Name, VBAToolKit.vtkTestPath & "\" & testModule1FileName
    vtkExportOneModule testWorkbook.VBProject, testModule2Name, VBAToolKit.vtkTestPath & "\" & testModule2FileName
    
    ' Get the size of the exported modules before the save
    oldSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    oldSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size
    
    ' Add the handler
    vtkAddBeforeSaveHandlerInDEVWorkbook testWorkbook

    ' Modify a module
    insertDummyProcedureInCodeModule testWorkbook.VBProject.VBComponents(testModule1Name).CodeModule, dummyInt:=1
    
    ' Save the workbook
    testWorkbook.Save
    
    ' Get the new size of the exported modules
    newSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule1FileName).Size
    newSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testModule2FileName).Size

    mAssert.Should oldSize1 < newSize1, testModule1FileName & " must be bigger after the save."
    mAssert.Should oldSize1 = newSize1, testModule2FileName & " must be of equal size after the save."

    On Error GoTo 0
    Exit Sub

Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook of Class Module vtkBeforeSaveHandlerTester"
    Resume Next
End Sub

Public Sub Test_vtkAddBeforeSaveHandlerInDEVWorkbook_noDEVInName()
    
    Dim error As Integer: error = -1
    
    On Error GoTo Test_vtkAddBeforeSaveHandlerInDevWorkbook_noDEVInName_Error
    
    ' Import relevant workbook from the templates folder
    Set existingProjectWorkbook = VBAToolKit.getTestFileFromTemplate(fileName:=existingProjectWorkbookName, openExcel:=True)

    ' Try to add the before save handler
    vtkAddBeforeSaveHandlerInDEVWorkbook existingProjectWorkbook

    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER, "The call must return a VTK_FORBIDDEN_PARAMETER error."

    On Error GoTo 0
    Exit Sub

Test_vtkAddBeforeSaveHandlerInDevWorkbook_noDEVInName_Error:
    error = Err.number
    Resume Next
    
End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkAddBeforeSaveHandlerInDEVWorkbook_noDEVInName"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook": Test_vtkAddBeforeSaveHandlerInDEVWorkbook_validWorkbook
        Case "Test_vtkAddBeforeSaveHandlerInDEVWorkbook_noDEVInName": Test_vtkAddBeforeSaveHandlerInDEVWorkbook_noDEVInName
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

