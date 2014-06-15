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
    mAssert.Equals Err.Number, VTK_NOTINITIALIZED, "Error sent by references"
    Set c = existingRefManager.allReferences
    mAssert.Equals Err.Number, VTK_NOTINITIALIZED, "Error sent by allReferences"
   On Error GoTo 0
   
End Sub

Public Sub TestInitializeStandardRefs()
'       Verify that standard references are added to the reference sheet at initialization
    Dim ws As Worksheet, r As Range
    Dim refNames(), i As Integer
    refNames = Array("Scripting", "VBIDE", "Shell32", "MSXML2", "ADODB", "VBAToolKit")
        
   On Error GoTo M_Error
    Set ws = newWorkBook.Sheets("vtkReferences")
    For i = LBound(refNames) To UBound(refNames)
        Set r = ws.Columns("A:A").Find(refNames(i), , xlValues, xlWhole)
        mAssert.Should Not r Is Nothing, "The reference named " & refNames(i) & " is not included in the reference sheet"
    Next i
   On Error GoTo 0
    Exit Sub
    
M_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description & "."
End Sub

Public Sub TestAddNewRefs()
'       Verify that new references are added to the reference sheet at initialization
    Dim ws As Worksheet, r As Range
    Dim refNames(), i As Integer
    refNames = Array("Scripting", "VBIDE", "Shell32", "MSXML2", "VBAToolKit", "EventSystemLib")
        
   On Error GoTo M_Error
    Set ws = existingWorkbook.Sheets("vtkReferences")
    For i = LBound(refNames) To UBound(refNames)
        Set r = ws.Columns("A:A").Find(refNames(i), , xlValues, xlWhole)
        mAssert.Should Not r Is Nothing, "The reference named " & refNames(i) & " is not included in the reference sheet"
    Next i
   On Error GoTo 0
    Exit Sub
    
M_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description & "."
End Sub

Public Sub TestAddNewRefWithRelativePath()
'       Verify that new references are added to the reference sheet at initialization
    Dim ws As Worksheet, r As Range
        
   On Error GoTo M_Error
    Set ws = existingWorkbook.Sheets("vtkReferences")
    Set r = ws.Columns("A:A").Find("VBAToolKit", , xlValues, xlWhole)
    mAssert.Equals ws.Cells(r.Row, 3), "Microsoft\AddIns\VBAToolKit.xlam", "The reference named VBAToolKit must be included with relative path in the reference sheet"
   On Error GoTo 0
    Exit Sub
    
M_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description & "."
End Sub

Public Sub TestGetAllReferencesFromNewWorkbook()
'       Verify that all references are listed with good values
    Dim ws As Worksheet, r As Range, c As Collection, ref As vtkReference, i As Integer
        
   On Error GoTo M_Error
    Set ws = newWorkBook.Sheets("vtkReferences")
    Set c = newRefManager.allReferences
    For i = 1 To c.Count
        Set ref = c(i)
        mAssert.Equals ref.id, "r" & i, "ID of the reference number " & i
        Set r = ws.Columns("A:A").Find(ref.name, , xlValues, xlWhole)
        mAssert.Equals ws.Cells(r.Row, 1), ref.name, "name of the reference number " & i
        mAssert.Equals ws.Cells(r.Row, 2), ref.GUID, "Guid of the reference number " & i
        If ref.GUID Like "" Then
            mAssert.Equals ws.Cells(r.Row, 3), ref.relPath, "Relative Path of the reference number " & i
        End If
    Next i
   On Error GoTo 0
    Exit Sub
    
M_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description & "."
End Sub

Public Sub TestGetAllReferencesFromExistingWorkbook()
'       Verify that all references are listed with good values
    Dim ws As Worksheet, r As Range, c As Collection, ref As vtkReference, i As Integer
        
   On Error GoTo M_Error
    Set ws = existingWorkbook.Sheets("vtkReferences")
    Set c = existingRefManager.allReferences
    For i = 1 To c.Count
        Set ref = c(i)
        mAssert.Equals ref.id, "r" & i, "ID of the reference number " & i
        Set r = ws.Columns("A:A").Find(ref.name, , xlValues, xlWhole)
        mAssert.Equals ws.Cells(r.Row, 1), ref.name, "name of the reference number " & i
        mAssert.Equals ws.Cells(r.Row, 2), ref.GUID, "Guid of the reference number " & i
        If ref.GUID Like "" Then
            mAssert.Equals ws.Cells(r.Row, 3), ref.relPath, "Relative Path of the reference number " & i
        End If
    Next i
   On Error GoTo 0
    Exit Sub
    
M_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description & "."
End Sub

Public Sub TestGetReferencesFromNewWorkbook()
'       Verify that all references are listed with good values
    Dim ws As Worksheet, r As Range, c As Collection, ref As vtkReference, i As Integer
        
   On Error GoTo M_Error
    Set ws = newWorkBook.Sheets("vtkReferences")
    
    ' Test references for the main configuration (#1)
    Set c = newRefManager.references(1)
    mAssert.Equals c.count, 10, "Reference count for the main configuration"
    mAssert.Equals c(1), "VBA", "Reference for the main configuration"
    mAssert.Equals c(2), "Excel", "Reference for the main configuration"
    mAssert.Equals c(3), "stdole", "Reference for the main configuration"
    mAssert.Equals c(4), "Office", "Reference for the main configuration"
    mAssert.Equals c(5), "MSForms", "Reference for the main configuration"
    mAssert.Equals c(6), "Scripting", "Reference for the main configuration"
    mAssert.Equals c(7), "VBIDE", "Reference for the main configuration"
    mAssert.Equals c(8), "Shell32", "Reference for the main configuration"
    mAssert.Equals c(9), "MSXML2", "Reference for the main configuration"
    mAssert.Equals c(10), "ADODB", "Reference for the main configuration"
    
    ' Test references for the development configuration (#2)
    Set c = newRefManager.references(2)
    mAssert.Equals c.count, 11, "Reference count for the development configuration"
    mAssert.Equals c(1), "VBA", "Reference for the development configuration"
    mAssert.Equals c(2), "Excel", "Reference for the development configuration"
    mAssert.Equals c(3), "stdole", "Reference for the development configuration"
    mAssert.Equals c(4), "Office", "Reference for the development configuration"
    mAssert.Equals c(5), "MSForms", "Reference for the development configuration"
    mAssert.Equals c(6), "Scripting", "Reference for the development configuration"
    mAssert.Equals c(7), "VBIDE", "Reference for the development configuration"
    mAssert.Equals c(8), "Shell32", "Reference for the development configuration"
    mAssert.Equals c(9), "MSXML2", "Reference for the development configuration"
    mAssert.Equals c(10), "ADODB", "Reference for the development configuration"
    mAssert.Equals c(11), "VBAToolKit", "Reference for the development configuration"
    
   On Error GoTo 0
    Exit Sub
   
M_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description & "."
End Sub

' Reste à tester :
'   récup des références (tester GUID et constitution du path absolu/relatif, new workbook et existing workbook)

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
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestInitializeStandardRefs"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddNewRefs"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddNewRefWithRelativePath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetAllReferencesFromNewWorkbook"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetAllReferencesFromExistingWorkbook"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetReferencesFromNewWorkbook"
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
        Case "TestInitializeStandardRefs": TestInitializeStandardRefs
        Case "TestAddNewRefs": TestAddNewRefs
        Case "TestAddNewRefWithRelativePath": TestAddNewRefWithRelativePath
        Case "TestGetAllReferencesFromNewWorkbook": TestGetAllReferencesFromNewWorkbook
        Case "TestGetAllReferencesFromExistingWorkbook": TestGetAllReferencesFromExistingWorkbook
        Case "TestGetReferencesFromNewWorkbook": TestGetReferencesFromNewWorkbook
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

