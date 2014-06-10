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

    Set mAssert = Assert
    
    Set newWorkBook = vtkCreateExcelWorkbookForTestWithProjectName("NewProject")    ' create an Excel project for tests
    Set newRefManager = New vtkReferenceManager
    Set newRefManager.Workbook = newWorkBook
    
    Set existingWorkbook = getTestFileFromTemplate(fileName:=existingWorkbookNameForTest, destinationName:=existingProjectName & "_DEV.xlsm", openExcel:=True)
    Set existingRefManager = New vtkReferenceManager
    Set existingRefManager.Workbook = existingWorkbook
End Sub

Private Sub ITestCase_TearDown()
    vtkCloseAndKillWorkbook Wb:=newWorkBook     ' close the new Excel project
    vtkCloseAndKillWorkbook Wb:=existingWorkbook ' close the existing Excel project
    Set newRefManager = Nothing             ' reset all reference managers
    Set existingRefManager = Nothing
End Sub

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestConfigurationSheetExistsInExistingWorkbook": TestConfigurationSheetExistsInExistingWorkbook
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

