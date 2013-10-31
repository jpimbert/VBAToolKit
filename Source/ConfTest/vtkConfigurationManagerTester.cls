VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfigurationManagerTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkConfigurationManagerTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the vtkConfigurationManager class
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

Private newWorkBook As Workbook         ' New Workbook created for each test
Private newWorkbookName As String
Private newConfManager As vtkConfigurationManager   ' Configuration Manager for the new workbook
Private newProjectName As String

Private existingWorkbook As Workbook    ' Existing Workbook opened for each test
Private Const existingWorkbookNameForTest As String = "WorkBookForConfigurationsTests.xlsm"
Private existingConfManager As vtkConfigurationManager   ' Configuration Manager for the new workbook
Private Const existingProjectName As String = "ExistingProject"
Private existingWorkbookName As String


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

End Sub

Private Sub ITestCase_TearDown()

End Sub

' TESTS TO WRITE
'
' The "load from file" functionality
'   - test that all the modules were loaded
'   - test that modules are well stored in function of the configuration they belong to
'   - test that all the references were loaded
'   - test that references are well stored in function of their configuration and their nature (guid or path)
'   - test that info was well loaded
'   - test that configurations were well loaded
'   - test that the above works also with a 1.0 file
'
' The "save at every modification" functionality
'   - when not enabled, file is not modified
'   - when enabled, file is modified
'
'
Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


