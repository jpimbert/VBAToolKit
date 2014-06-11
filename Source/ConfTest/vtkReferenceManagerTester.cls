VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkReferenceManagerTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkReferenceManagerTester11
' Author    : Jean-Pierre Imbert
' Date      : 10/06/2014
' Purpose   : Test the vtkReferenceManager class
'
' Copyright 2014 Skwal-Soft (http://skwalsoft.com)
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

Private newWorkBook As Workbook         ' New Workbook created for each test
Private newRefManager As vtkReferenceManager   ' Reference Manager for the new workbook

Private existingWorkbook As Workbook    ' Existing Workbook opened for each test
Private Const existingWorkbookNameForTest As String = "WorkBookForConfigurationsTests.xlsm"
Private existingRefManager As vtkReferenceManager   ' Reference Manager for the existing workbook

Private Sub SaveThenReOpenNewWorkbook()
' Utility Function for saving then re-open the new workbook and re-initialize the reference manager
    Dim savedNewWorkbookPath As String
    savedNewWorkbookPath = newWorkBook.FullName
'    If newWorkbook.path = "" Then
'        newWorkbook.Close SaveChanges:=True, Filename:=savedNewWorkbookPath
'       Else
'        newWorkbook.Close SaveChanges:=True
'    End If
    newWorkBook.Close saveChanges:=True
    Set newWorkBook = Workbooks.Open(fileName:=savedNewWorkbookPath)
    Set newRefManager = New vtkReferenceManager
    Set newRefManager.Workbook = newWorkBook
End Sub

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
    Dim newConfManager As vtkConfigurationManager

    Set mAssert = Assert
    
    Set newWorkBook = vtkCreateExcelWorkbookForTestWithProjectName("NewProject")    ' create an Excel project for tests
    Set newConfManager = vtkConfigurationManagerForProject(newWorkBook.VBProject.name) ' For vtkConfigurations sheet initialization
    Set newRefManager = New vtkReferenceManager
    newRefManager.init Wb:=newWorkBook, confCount:=2, nbTitleColumnsInConfSheet:=1
    
    Set existingWorkbook = getTestFileFromTemplate(fileName:=existingWorkbookNameForTest, destinationName:="ExistingProject_DEV.xlsm", openExcel:=True)
    Set existingRefManager = New vtkReferenceManager
    existingRefManager.init Wb:=existingWorkbook, confCount:=2, nbTitleColumnsInConfSheet:=1

End Sub

Private Sub ITestCase_TearDown()
    vtkCloseAndKillWorkbook Wb:=newWorkBook     ' close the new Excel project
    vtkCloseAndKillWorkbook Wb:=existingWorkbook ' close the existing Excel project
    vtkResetConfigurationManagers             ' reset all configuration managers
    Set newRefManager = Nothing
    Set existingRefManager = Nothing
End Sub

Public Sub TestReferenceSheetCreationForNewProject()
'       Verify that a Reference Sheet is created in a new project
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = newWorkBook.Sheets("vtkReferences")
    mAssert.Equals Err, 0, "Error number : A reference manager must create a Reference sheet"
    mAssert.Should ws Is newRefManager.referenceSheet, "The referenceSheet property of the ref manager must be equal to the reference sheet of the workbook"
    On Error GoTo 0
End Sub

Public Sub TestReferenceSheetRetrievalInExistingProject()
'       Verify that the Reference Sheet of existing project is retrieved
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = existingWorkbook.Sheets("vtkReferences")
    mAssert.Equals Err, 0, "Error number : A reference sheet must exist in the existing workbook"
    mAssert.Should ws Is existingRefManager.referenceSheet, "The referenceSheet property of the ref manager must be equal to the reference sheet of the workbook"
    On Error GoTo 0
End Sub

Public Sub TestReferenceSheetFormatForNewProjet()
'       Verify the newly created Reference sheet format of a new project
    Dim ws As Worksheet
   On Error GoTo TestReferenceSheetFormatForNewProjet_Error

    Set ws = newRefManager.referenceSheet
    mAssert.Equals ws.Range("A1"), "vtkReferences v1.0", "Expected identification of the reference sheet"
    mAssert.Equals ws.Range("B1"), "", "Expected Title for GUID column"
    mAssert.Equals ws.Range("C1"), "-", "Expected Title for FullPath column"
    mAssert.Equals ws.Range("D1"), newWorkBook.VBProject.name, "Expected Title for main project column"
    mAssert.Equals ws.Range("E1"), newWorkBook.VBProject.name & "_DEV", "Expected Title for DEV project column"

   On Error GoTo 0
   Exit Sub

TestReferenceSheetFormatForNewProjet_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description & "."
End Sub

Public Sub TestReferenceSheetFormatForExistingProjet()
'       Verify the retrieved Reference sheet format of an existing project
    Dim ws As Worksheet
   On Error GoTo TestReferenceSheetFormatForExistingProjet_Error

    Set ws = existingRefManager.referenceSheet
    mAssert.Equals ws.Range("A1"), "vtkReferences v1.0", "Expected identification of the reference sheet"
    mAssert.Equals ws.Range("B1"), "", "Expected Title for GUID column"
    mAssert.Equals ws.Range("C1"), "-", "Expected Title for FullPath column"
    mAssert.Equals ws.Range("D1"), existingWorkbook.VBProject.name, "Expected Title for main project column"
    mAssert.Equals ws.Range("E1"), existingWorkbook.VBProject.name & "_DEV", "Expected Title for DEV project column"

   On Error GoTo 0
   Exit Sub

TestReferenceSheetFormatForExistingProjet_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description & "."
End Sub

Public Sub TestAlignColumnCount()
'       Verify that the retrieved Reference sheet is incremented with new configurations if necessary
    Dim ws As Worksheet, confManager As vtkConfigurationManager
    
    ' Add a new configuration in the existing project
   On Error GoTo TestAlignColumnCount_Error

    Set confManager = vtkConfigurationManagerForProject("ExistingProject")
    confManager.addConfiguration "NewConfiguration"
    
    ' Reinitialize the referenceManager
    existingRefManager.init Nothing, 0, nbTitleColumnsInConfSheet:=1
    existingRefManager.init existingWorkbook, 3, nbTitleColumnsInConfSheet:=1
    
    Set ws = existingRefManager.referenceSheet
    mAssert.Equals ws.Range("A1"), "vtkReferences v1.0", "Expected identification of the reference sheet"
    mAssert.Equals ws.Range("B1"), "", "Expected Title for GUID column"
    mAssert.Equals ws.Range("C1"), "-", "Expected Title for FullPath column"
    mAssert.Equals ws.Range("D1"), existingWorkbook.VBProject.name, "Expected Title for main project column"
    mAssert.Equals ws.Range("E1"), existingWorkbook.VBProject.name & "_DEV", "Expected Title for DEV project column"
    mAssert.Equals ws.Range("F1"), "NewConfiguration", "Expected Title for the new configuration column"
    
    Set confManager = Nothing

   On Error GoTo 0
   Exit Sub

TestAlignColumnCount_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description & "."
End Sub

Public Sub TestAddConfigurationNotInitialized()
'       Verify that addConfiguration method send an error when called while the instance is not initialized
    
    ' Reinit the reference manager
    existingRefManager.init Nothing, 0, nbTitleColumnsInConfSheet:=1
    
   On Error Resume Next
    existingRefManager.addConfiguration
    mAssert.Equals Err.Number, VTK_NOTINITIALIZED, "Error sent by addConfiguration"
   On Error GoTo 0
   
End Sub

Public Sub TestAddConfiguration()
'       Verify the addConfiguration method in nominal case
    Dim ws As Worksheet, confManager As vtkConfigurationManager
    
    ' Add a new configuration in the existing project
   On Error GoTo TestAddConfiguration_Error

    Set confManager = vtkConfigurationManagerForProject("ExistingProject")
    confManager.addConfiguration "NewConfiguration"
    
    ' add the configuration to the reference sheet (this call will be implied in a future release of cofManager
    existingRefManager.addConfiguration
    
    Set ws = existingRefManager.referenceSheet
    mAssert.Equals ws.Range("A1"), "vtkReferences v1.0", "Expected identification of the reference sheet"
    mAssert.Equals ws.Range("B1"), "", "Expected Title for GUID column"
    mAssert.Equals ws.Range("C1"), "-", "Expected Title for FullPath column"
    mAssert.Equals ws.Range("D1"), existingWorkbook.VBProject.name, "Expected Title for main project column"
    mAssert.Equals ws.Range("E1"), existingWorkbook.VBProject.name & "_DEV", "Expected Title for DEV project column"
    mAssert.Equals ws.Range("F1"), "NewConfiguration", "Expected Title for the new configuration column"
    
    Set confManager = Nothing

   On Error GoTo 0
   Exit Sub

TestAddConfiguration_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description & "."
End Sub

Public Sub TestReferencesNotInitialized()
'       Verify that references method send an error when called while the instance is not initialized
    Dim c As Collection
    
    ' Reinit the reference manager
    existingRefManager.init Nothing, 0, nbTitleColumnsInConfSheet:=1
    
   On Error Resume Next
    Set c = existingRefManager.references(1)
    mAssert.Equals Err.Number, VTK_NOTINITIALIZED, "Error sent by addConfiguration"
   On Error GoTo 0
   
End Sub

' a tester AddConfiguration sur un refManager non initialisé (doit remonter une erreur ??)

' tester l'appel à references alors que l'instance n'est pas initialisée (VTK_NOTINITIALIZED)

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferenceSheetCreationForNewProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferenceSheetRetrievalInExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferenceSheetFormatForNewProjet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferenceSheetFormatForExistingProjet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAlignColumnCount"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddConfigurationNotInitialized"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddConfiguration"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferencesNotInitialized"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestReferenceSheetCreationForNewProject": TestReferenceSheetCreationForNewProject
        Case "TestReferenceSheetRetrievalInExistingProject": TestReferenceSheetRetrievalInExistingProject
        Case "TestReferenceSheetFormatForNewProjet": TestReferenceSheetFormatForNewProjet
        Case "TestReferenceSheetFormatForExistingProjet": TestReferenceSheetFormatForExistingProjet
        Case "TestAlignColumnCount": TestAlignColumnCount
        Case "TestAddConfigurationNotInitialized": TestAddConfigurationNotInitialized
        Case "TestAddConfiguration": TestAddConfiguration
        Case "TestReferencesNotInitialized": TestReferencesNotInitialized
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

