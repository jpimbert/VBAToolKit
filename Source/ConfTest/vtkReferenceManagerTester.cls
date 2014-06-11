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
    newRefManager.init Wb:=newWorkBook, confCount:=2
    
    Set existingWorkbook = getTestFileFromTemplate(fileName:=existingWorkbookNameForTest, destinationName:="ExistingProject_DEV.xlsm", openExcel:=True)
    Set existingRefManager = New vtkReferenceManager
    existingRefManager.init Wb:=existingWorkbook, confCount:=2
End Sub

Private Sub ITestCase_TearDown()
    vtkCloseAndKillWorkbook Wb:=newWorkBook     ' close the new Excel project
    vtkCloseAndKillWorkbook Wb:=existingWorkbook ' close the existing Excel project
    Set newRefManager = Nothing             ' reset all reference managers
    Set existingRefManager = Nothing
End Sub

Public Sub TestReferenceSheetCreationForNewProject()
'       Verify that a Configuration Sheet is created in a new project
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = newWorkBook.Sheets("vtkReferences")
    mAssert.Equals Err, 0, "Error number : A reference manager must create a Reference sheet"
    mAssert.Should ws Is newRefManager.referenceSheet, "The referenceSheet property of the ref manager must be equal to the reference sheet of the workbook"
    On Error GoTo 0
End Sub

Public Sub TestReferenceSheetRetrievalInExistingProject()
'       Verify that a Configuration Sheet is created in a new project
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = existingWorkbook.Sheets("vtkReferences")
    mAssert.Equals Err, 0, "Error number : A reference sheet must exist in the existing workbook"
    mAssert.Should ws Is existingRefManager.referenceSheet, "The referenceSheet property of the ref manager must be equal to the reference sheet of the workbook"
    On Error GoTo 0
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferenceSheetCreationForNewProject"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestReferenceSheetCreationForNewProject": TestReferenceSheetCreationForNewProject
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

