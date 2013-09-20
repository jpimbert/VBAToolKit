VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkTestUtilitiesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkTestUtilitiesTester
' Author    : Champonnois
' Date      : 17/09/2013
' Purpose   : Test the TestUtilities functions
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

Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

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
    vtkTestUtilities.prepare ThisWorkbook
End Sub

Private Sub ITestCase_TearDown()
    vtkCleanFolder VBAToolKit.vtkTestPath
End Sub

Public Sub Test_vtkTestPath()
    ' get the path of a project that is not the active one
    Dim wb1 As Workbook, wb2 As Workbook, p1 As vtkProject, p2 As vtkProject
    
    On Error GoTo Test_vtkTestPath_Error
    
    Set p1 = vtkProjectForName("TestProject1")
    Set p2 = vtkProjectForName("TestProject2")
    vtkCreateProject path:=VBAToolKit.vtkTestPath, name:="TestProject1", displayError:=False
    Set wb1 = Workbooks(p1.workbookDEVName)
    vtkCreateProject path:=VBAToolKit.vtkTestPath, name:="TestProject2", displayError:=False
    Set wb2 = Workbooks(p2.workbookDEVName)
    
    wb1.Activate
    vtkTestUtilities.prepare ActiveWorkbook
    wb2.Activate
    
    mAssert.Should (vtkTestPath Like "*TestProject1\Tests"), "Bad Test path for TestProject1 (" & vtkTestPath & "), should end with TestProject1\Tests"
    
    'Close file Excel and deleting the project "TestProject1" and "TestProject2"
    wb1.Close (False)
    wb2.Close (False)
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkTestPath_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkTestPath of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

Public Sub Test_vtkGetFileFromTemplate()
    'Verify the copy file from template in the folder test
    Dim wbb As Workbook
    
    On Error GoTo Test_vtkGetFileFromTemplate_Error
    
    Set wbb = getTestFileFromTemplate(fileName:="TestModule.bas")
    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "TestModule.bas"), "TestModule.bas", "Bad test the file TestModule hasn't been copied in folder Test"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetFileFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetFileFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

Public Sub Test_vtkGetNotFileFromTemplate()
    'Verifying an error is displayed if one doesn't possess the file template
    Dim wbb As Workbook
    
    On Error GoTo Test_vtkGetNotFileFromTemplate_Error
    
    Set wbb = getTestFileFromTemplate(fileName:="Error.bas")
    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "Error.bas"), "", "Bad test the file Error doesn't exist from template"
    On Error GoTo 0
    Exit Sub
Test_vtkGetNotFileFromTemplate_Error:
    If Err.number <> VTK_FILE_NOT_FOUND Then mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetNotFileFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

Public Sub Test_vtkGetReplaceFileFromTemplate()
    Dim fso As FileSystemObject, wbb As Workbook

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo Test_vtkGetReplaceFileFromTemplate_Error
    
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestModule.bas")
    
    Set wbb = getTestFileFromTemplate(fileName:="TestModule.bas")
    
    mAssert.Should vtkTextFileReader(VBAToolKit.vtkTestPath & "\" & "TestModule.bas") <> "", "the copied file doesn't replace the existing file"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetReplaceFileFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetReplaceFileFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkTestPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetNotFileFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetReplaceFileFromTemplate"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkTestPath": Test_vtkTestPath
        Case "Test_vtkGetFileFromTemplate": Test_vtkGetFileFromTemplate
        Case "Test_vtkGetNotFileFromTemplate": Test_vtkGetNotFileFromTemplate
        Case "Test_vtkGetReplaceFileFromTemplate": Test_vtkGetReplaceFileFromTemplate
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

