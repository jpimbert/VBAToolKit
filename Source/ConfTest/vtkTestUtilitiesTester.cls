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


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkTestPath
' Author    : Champonnois
' Purpose   : Get the path of a project that is not the active one
'---------------------------------------------------------------------------------------
Public Sub Test_vtkTestPath()

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


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileFromTemplate
' Author    : Champonnois
' Purpose   : Verify the copy file from template in the folder test
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileFromTemplate()

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


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetNotFileFromTemplate
' Author    : Champonnois
' Purpose   : Verifying an error is displayed if one doesn't possess the file template
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetNotFileFromTemplate()

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


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetReplaceFileFromTemplate
' Author    : Champonnois
' Purpose   : Test the replacement of an existing file
'---------------------------------------------------------------------------------------
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


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileFromTemplateOpenExcel
' Author    : Champonnois
' Purpose   : After copying an Excel file, verify if the file is opened
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileFromTemplateOpenExcel()

    Dim wbb As Workbook

    On Error GoTo Test_vtkGetFileFromTemplateOpenExcel_Error

    Set wbb = getTestFileFromTemplate(fileName:="UninitializedProject.xlsm", openExcel:=True)

    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "UninitializedProject.xlsm"), "UninitializedProject.xlsm", "Bad test the file Error doesn't exist from template"
    mAssert.Equals wbb.path, VBAToolKit.vtkTestPath, "the Excel file isn't open in the right directory"
    mAssert.Should VtkWorkbookIsOpen(wbb.name), "The file Excel isn't open"

    wbb.Close (False)

    On Error GoTo 0

    Exit Sub

Test_vtkGetFileFromTemplateOpenExcel_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetFileFromTemplateOpenExcel of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileFromTemplateExcelNotOpen
' Author    : Champonnois
' Purpose   : After copying an Excel file, verify if the file isn't opened
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileFromTemplateExcelNotOpen()

    Dim wbb As Workbook

    On Error GoTo Test_vtkGetFileFromTemplateExcelNotOpen_Error

    Set wbb = getTestFileFromTemplate(fileName:="UninitializedProject.xlsm", openExcel:=False)

    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "UninitializedProject.xlsm"), "UninitializedProject.xlsm", "Bad test the file Error doesn't exist from template"
    mAssert.Should wbb Is Nothing, "The workbook shouldn't  exist"
    mAssert.Should Not VtkWorkbookIsOpen("UninitializedProject.xlsm"), "the file Excel shouldn't be open"

    On Error GoTo 0

    Exit Sub

Test_vtkGetFileFromTemplateExcelNotOpen_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetFileFromTemplateExcelNotOpen of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetErrorFolderFromTemplate
' Author    : Champonnois
' Purpose   : Verify the copy file from template in the folder test
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetErrorFolderFromTemplate()
    
    Dim wbb As Workbook

    On Error GoTo Test_vtkGetErrorFolderFromTemplate_Error
    On Error Resume Next

    Set wbb = getTestFileFromTemplate(fileName:="Folder")

    mAssert.Should wbb Is Nothing, "The workbook shouldn't  exist"
    mAssert.Equals Err.number, VTK_DOESNT_COPY_FOLDER, "The proper error must be raised when trying to copy a folder"
    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "Folder"), "", "The folder was copy"
    
    On Error GoTo 0
    Exit Sub

Test_vtkGetErrorFolderFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetErrorFolderFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileNotVisibleFromTemplate
' Author    : Champonnois
' Purpose   : Verify the copy a hidden file from template
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileNotVisibleFromTemplate()

    Dim wbb As Workbook

    On Error GoTo Test_vtkGetFileNotVisibleFromTemplate_Error

    Set wbb = getTestFileFromTemplate(fileName:="FileNotVisible.bas")

    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "FileNotVisible.bas"), "FileNotVisible.bas", "Bad test the file Error doesn't exist from template"

    On Error GoTo 0
    Exit Sub

Test_vtkGetFileNotVisibleFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetFileNotVisibleFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileFromTemplate_ParamDestination
' Author    : Champonnois
' Purpose   : Verify the copy file from template by providing the method the copy path
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileFromTemplate_ParamDestination()

    Dim wbb As Workbook

    On Error GoTo Test_vtkGetFileFromTemplate_ParamDestination_Error

    Set wbb = getTestFileFromTemplate(fileName:="TestModule.bas", destinationName:="TestParamDestinationName.bas")
    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "TestParamDestinationName.bas"), "TestParamDestinationName.bas", "Bad test the file TestModule hasn't been copied in folder Test"

    On Error GoTo 0
    Exit Sub

Test_vtkGetFileFromTemplate_ParamDestination_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetFileFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetNotFileFromTemplate_ParamDestination
' Author    : Champonnois
' Purpose   : Same procedure as the copy of a nonexistent file, but with the given path
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetNotFileFromTemplate_ParamDestination()
    '
    Dim wbb As Workbook

    On Error GoTo Test_vtkGetNotFileFromTemplate_ParamDestination_Error
    On Error Resume Next
    
    Set wbb = getTestFileFromTemplate(fileName:="Error.bas", destinationName:="ErrorParamDestinationName.bas")
    
    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "ErrorParamDestinationName.bas"), "", "Bad test the file Error doesn't exist from template"
    mAssert.Equals Err.number, VTK_FILE_NOT_FOUND, "The proper error must be raised when trying to copy a folder"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetNotFileFromTemplate_ParamDestination_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetNotFileFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetReplaceFileFromTemplate_ParamDestination
' Author    : Champonnois
' Purpose   : Copy of a file in the template generates the replacement of an existing file in the given path
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetReplaceFileFromTemplate_ParamDestination()
    
    Dim fso As FileSystemObject, wbb As Workbook, FileLenBeforeCopy As Long, FileLenAfterCopy As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Test_vtkGetReplaceFileFromTemplate_ParamDestination_Error

    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestParamDestinationName.bas")

    FileLenBeforeCopy = FileLen(VBAToolKit.vtkTestPath & "\" & "TestParamDestinationName.bas")

    Set wbb = getTestFileFromTemplate(fileName:="TestModule.bas", destinationName:="TestParamDestinationName.bas")

    FileLenAfterCopy = FileLen(VBAToolKit.vtkTestPath & "\" & "TestParamDestinationName.bas")

    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "TestParamDestinationName.bas"), "TestParamDestinationName.bas", "Bad test the file Error doesn't exist from template"
    mAssert.Should FileLenBeforeCopy < FileLenAfterCopy, "the copied file doesn't replace the existing file"

    On Error GoTo 0
    Exit Sub

Test_vtkGetReplaceFileFromTemplate_ParamDestination_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetReplaceFileFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verify the copy folder from template in the folder test
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderFromTemplate()
  
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetNotFolderFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verifying an error is displayed if one doesn't possess the folder template
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetNotFolderFromTemplate()
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetErrorFileFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetErrorFileFromTemplate()
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetReplaceFolderFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Test the replacement of an existing folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetReplaceFolderFromTemplate()
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderFromTemplate_ParamDestination
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verify the copy folder from template by providing the method the copy path
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderFromTemplate_ParamDestination()

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetNotFolderFromTemplate_ParamDestination
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Same procedure as the copy of a nonexistent folder, but with the given path
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetNotFolderFromTemplate_ParamDestination()
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetReplaceFolderFromTemplate_ParamDestination
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Copy of a folder in the template generates the replacement of an existing file in the given path
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetReplaceFolderFromTemplate_ParamDestination()
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderNotVisibleFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verify the copy a hidden file from template
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderNotVisibleFromTemplate()
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderNotVisibleInFolderFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verify the copy a hidden folder from template
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderNotVisibleInFolderFromTemplate()
   
End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkTestPath"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetNotFileFromTemplate"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetReplaceFileFromTemplate"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplateOpenExcel"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplateExcelNotOpen"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetErrorFolderFromTemplate"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileNotVisibleFromTemplate"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate_ParamDestination"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetNotFileFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetReplaceFileFromTemplate_ParamDestination"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetNotFolderFromTemplate"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetErrorFileFromTemplate"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetReplaceFolderFromTemplate"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate_ParamDestination"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetNotFolderFromTemplate_ParamDestination"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetReplaceFolderFromTemplate_ParamDestination"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderNotVisibleFromTemplate"
    'ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderNotVisibleInFolderFromTemplate"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkTestPath": Test_vtkTestPath
        Case "Test_vtkGetFileFromTemplate": Test_vtkGetFileFromTemplate
        Case "Test_vtkGetNotFileFromTemplate": Test_vtkGetNotFileFromTemplate
        Case "Test_vtkGetReplaceFileFromTemplate": Test_vtkGetReplaceFileFromTemplate
        Case "Test_vtkGetFileFromTemplateOpenExcel": Test_vtkGetFileFromTemplateOpenExcel
        Case "Test_vtkGetFileFromTemplateExcelNotOpen": Test_vtkGetFileFromTemplateExcelNotOpen
        Case "Test_vtkGetErrorFolderFromTemplate": Test_vtkGetErrorFolderFromTemplate
        Case "Test_vtkGetFileNotVisibleFromTemplate": Test_vtkGetFileNotVisibleFromTemplate
        Case "Test_vtkGetFileFromTemplate_ParamDestination": Test_vtkGetFileFromTemplate_ParamDestination
        Case "Test_vtkGetNotFileFromTemplate_ParamDestination": Test_vtkGetNotFileFromTemplate_ParamDestination
        Case "Test_vtkGetReplaceFileFromTemplate_ParamDestination": Test_vtkGetReplaceFileFromTemplate_ParamDestination
        Case "Test_vtkGetFolderFromTemplate": Test_vtkGetFolderFromTemplate
        Case "Test_vtkGetNotFolderFromTemplate": Test_vtkGetNotFolderFromTemplate
        Case "Test_vtkGetErrorFileFromTemplate": Test_vtkGetErrorFileFromTemplate
        Case "Test_vtkGetReplaceFolderFromTemplate": Test_vtkGetReplaceFolderFromTemplate
        Case "Test_vtkGetFolderFromTemplate_ParamDestination": Test_vtkGetFolderFromTemplate_ParamDestination
        Case "Test_vtkGetNotFolderFromTemplate_ParamDestination": Test_vtkGetNotFolderFromTemplate_ParamDestination
        Case "Test_vtkGetReplaceFolderFromTemplate_ParamDestination": Test_vtkGetReplaceFolderFromTemplate_ParamDestination
        Case "Test_vtkGetFolderNotVisibleFromTemplate": Test_vtkGetFolderNotVisibleFromTemplate
        Case "Test_vtkGetFolderNotVisibleInFolderFromTemplate": Test_vtkGetFolderNotVisibleInFolderFromTemplate
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

