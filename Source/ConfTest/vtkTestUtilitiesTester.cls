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
    vtkCreateProject Path:=VBAToolKit.vtkTestPath, name:="TestProject1", displayError:=False
    Set wb1 = Workbooks(p1.workbookDEVName)
    vtkCreateProject Path:=VBAToolKit.vtkTestPath, name:="TestProject2", displayError:=False
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
    mAssert.Equals wbb.Path, VBAToolKit.vtkTestPath, "the Excel file isn't open in the right directory"
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
' Purpose   : Same procedure as the copy of a file, but with a new name for the copied file
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
' Purpose   : Same procedure as the copy of a nonexistent file, but with a new name for the copied file
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
' Purpose   : Same procedure as the replace of a file, but with a new name for the copied file
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
    Dim fso As FileSystemObject
    
    On Error GoTo Test_vtkGetFolderFromTemplate_Error
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    getTestFolderFromTemplate folderName:="Folder"
    
    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Folder"), "Bad test the folder Folder hasn't been copied in folder Test"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetFolderFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetFolderFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetNotFolderFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verifying an error is displayed if one doesn't possess the folder in template
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetNotFolderFromTemplate()
    Dim fso As FileSystemObject
    
    On Error GoTo Test_vtkGetNotFolderFromTemplate_Error
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    getTestFolderFromTemplate folderName:="Error"
    
    mAssert.Equals fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Error"), False, "Bad test the folder Folder hasn't been copied in folder Test"
    mAssert.Equals Err.number, VTK_FOLDER_NOT_FOUND, "The proper error must be raised when trying to copy a folder"
    
    On Error GoTo 0
    
    Exit Sub
Test_vtkGetNotFolderFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetNotFolderFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetErrorFileFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verifying that you can't copy the file and displays an error
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetErrorFileFromTemplate()

    Dim fso As FileSystemObject
    
    On Error GoTo Test_vtkGetErrorFileFromTemplate_Error
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    getTestFolderFromTemplate folderName:="Module.bas"

    mAssert.Equals fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Module.bas"), False, "Bad test the folder Folder hasn't been copied in folder Test"
    mAssert.Equals Err.number, VTK_FOLDER_NOT_FOUND, "The proper error must be raised when trying to copy a folder"
    
    On Error GoTo 0
    Exit Sub
Test_vtkGetErrorFileFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetErrorFolderFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetReplaceFolderFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Test the replacement of an existing folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetReplaceFolderFromTemplate()
 
    Dim fso As FileSystemObject, FolderBeforeCopy As Folder, FolderAfterCopy As Folder, SizeFolderBeforeCopy As Integer, SizeFolderAfterCopy As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo Test_vtkGetReplaceFolderFromTemplate_Error
    
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "Folder")
    
    Set FolderBeforeCopy = fso.GetFolder(VBAToolKit.vtkTestPath & "\" & "Folder")
    SizeFolderBeforeCopy = FolderBeforeCopy.Size
    
    getTestFolderFromTemplate folderName:="Folder"
    
    Set FolderAfterCopy = fso.GetFolder(VBAToolKit.vtkTestPath & "\" & "Folder")
    SizeFolderAfterCopy = FolderAfterCopy.Size
    
    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Folder"), "the copied file doesn't replace the existing file"
    mAssert.Should SizeFolderBeforeCopy < SizeFolderAfterCopy, "the copied file doesn't replace the existing file"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetReplaceFolderFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetReplaceFolderFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderFromTemplate_ParamDestination
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Same procedure as the copy of a folder, but with a new name for the copied folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderFromTemplate_ParamDestination()

    Dim fso As FileSystemObject
    
    On Error GoTo Test_vtkGetFolderFromTemplate_ParamDestination_Error
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    getTestFolderFromTemplate folderName:="Folder", destinationName:="TestRenameFolder"
    
    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & "TestRenameFolder"), "Bad test the folder Folder hasn't been copied in folder Test"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetFolderFromTemplate_ParamDestination_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetFolderFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetNotFolderFromTemplate_ParamDestination
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Same procedure as the copy of a nonexistent folder, but with a new name for the copied folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetNotFolderFromTemplate_ParamDestination()

    Dim fso As FileSystemObject
    
    On Error GoTo Test_vtkGetNotFolderFromTemplate_ParamDestination_Error
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    getTestFolderFromTemplate folderName:="Error", destinationName:="TestRenameError"
    
    mAssert.Equals fso.folderExists(VBAToolKit.vtkTestPath & "\" & "TestRenameError"), False, "Bad test the folder Folder hasn't been copied in folder Test"
    mAssert.Equals Err.number, VTK_FOLDER_NOT_FOUND, "The proper error must be raised when trying to copy a folder"
    
    On Error GoTo 0
    
    Exit Sub
Test_vtkGetNotFolderFromTemplate_ParamDestination_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetNotFolderFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetReplaceFolderFromTemplate_ParamDestination
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Same procedure as the replace of a folder, but with a new name for the copied folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetReplaceFolderFromTemplate_ParamDestination()

    Dim fso As FileSystemObject, FolderBeforeCopy As Folder, FolderAfterCopy As Folder, SizeFolderBeforeCopy As Integer, SizeFolderAfterCopy As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo Test_vtkGetReplaceFolderFromTemplate_ParamDestination_Error
    
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "TestRenameFolder")
    
    Set FolderBeforeCopy = fso.GetFolder(VBAToolKit.vtkTestPath & "\" & "TestRenameFolder")
    SizeFolderBeforeCopy = FolderBeforeCopy.Size
    
    getTestFolderFromTemplate folderName:="Folder", destinationName:="TestRenameFolder"
    
    Set FolderAfterCopy = fso.GetFolder(VBAToolKit.vtkTestPath & "\" & "TestRenameFolder")
    SizeFolderAfterCopy = FolderAfterCopy.Size
    
    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & "TestRenameFolder"), "the copied file doesn't replace the existing file"
    mAssert.Should SizeFolderBeforeCopy < SizeFolderAfterCopy, "the copied file doesn't replace the existing file"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetReplaceFolderFromTemplate_ParamDestination_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetReplaceFolderFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderNotVisibleFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verify the copy a hidden folder from template
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderNotVisibleFromTemplate()

    Dim fso As FileSystemObject

    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Test_vtkGetFolderNotVisibleFromTemplate_Error

    getTestFolderFromTemplate folderName:="FolderNotVisible"

    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & "FolderNotVisible"), "the copied file doesn't replace the existing file"

    On Error GoTo 0
    Exit Sub

Test_vtkGetFolderNotVisibleFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetFolderNotVisibleFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderNotVisibleInFolderFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verify the copy a hidden subfolder from template
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderNotVisibleInFolderFromTemplate()

    Dim fso As FileSystemObject

    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Test_vtkGetFolderNotVisibleInFolderFromTemplate_Error

    getTestFolderFromTemplate folderName:="Folder2"

    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Folder2"), "the copied file doesn't replace the existing file"

    On Error GoTo 0
    Exit Sub

Test_vtkGetFolderNotVisibleInFolderFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetFolderNotVisibleInFolderFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderCompleteTreeFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verification that the entire tree checked is copied
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderCompleteTreeFromTemplate()

    Dim fso As FileSystemObject

    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Test_vtkGetFolderCompleteTreeFromTemplate_Error

    getTestFolderFromTemplate folderName:="Folder"

    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Folder"), "the copied file doesn't replace the existing file"
    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & "Folder" & "\" & "TestClass.cls"), "the copied file doesn't replace the existing file"
    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Folder" & "\" & "SubFolder"), "the copied file doesn't replace the existing file"
    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & "Folder" & "\" & "SubFolder" & "\" & "TestModule.bas"), "the copied file doesn't replace the existing file"
    
    On Error GoTo 0
    Exit Sub

Test_vtkGetFolderCompleteTreeFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGetFolderCompleteTreeFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkResetTestFolder
' Author    : Champonnois
' Date      : 25/09/2013
' Purpose   : remove the file in Test folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkResetTestFolder()

    Dim fso As FileSystemObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Test_vtkResetTestFolder_Error

    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestResetFile.bas")
    
    resetTestFolder
    
    mAssert.Equals fso.FileExists(VBAToolKit.vtkTestPath & "\" & "TestResetFile.bas"), False, "The file TestResetFile hasn't been deleted"

    On Error GoTo 0
    Exit Sub

Test_vtkResetTestFolder_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkResetTestFolder of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkResetTestFolderFileOpen
' Author    : Champonnois
' Date      : 25/09/2013
' Purpose   : remove the file that is open in Test folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkResetTestFolderFileOpen()

    Dim fso As FileSystemObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Test_vtkResetTestFolderFileOpen_Error

    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestResetFileOpen.txt")
    fso.OpenTextFile (VBAToolKit.vtkTestPath & "\" & "TestResetFileOpen.txt")
   
    resetTestFolder
    
    mAssert.Equals fso.FileExists(VBAToolKit.vtkTestPath & "\" & "TestResetFileOpen.txt"), False, "The file TestResetFileOpen hasn't been deleted or the file is open"

    On Error GoTo 0
    Exit Sub

Test_vtkResetTestFolderFileOpen_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkResetTestFolderFileOpen of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkResetTestFolderEmptyFolder
' Author    : Champonnois
' Date      : 25/09/2013
' Purpose   : remove the empty folder in Test folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkResetTestFolderEmptyFolder()

    Dim fso As FileSystemObject, folderTest As Folder
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Test_vtkResetTestFolderEmptyFolder_Error
    
    resetTestFolder
    
    Set folderTest = fso.GetFolder(VBAToolKit.vtkTestPath)
    
    mAssert.Equals folderTest.Files.count, 0, "All files hasn't been deleted"
    mAssert.Equals folderTest.SubFolders.count, 0, "All folders hasn't been deleted"

    On Error GoTo 0
    Exit Sub

Test_vtkResetTestFolderEmptyFolder_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkResetTestFolderEmptyFolder of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkResetTestFolderContainsFiles
' Author    : Champonnois
' Date      : 25/09/2013
' Purpose   : remove the Test folder contains files
'---------------------------------------------------------------------------------------
Public Sub Test_vtkResetTestFolderContainsFiles()

    Dim fso As FileSystemObject, folderTest As Folder
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Test_vtkResetTestFolderContainsFiles_Error
    
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestTexte.txt")
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestBas.bas")
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestCls.cls")
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestXlsm.xlsm")
    
    resetTestFolder
    
    Set folderTest = fso.GetFolder(VBAToolKit.vtkTestPath)
    
    mAssert.Equals folderTest.Files.count, 0, "All files hasn't been deleted"

    On Error GoTo 0
    Exit Sub

Test_vtkResetTestFolderContainsFiles_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkResetTestFolderContainsFiles of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkResetTestFolderContainsFolderAndFiles
' Author    : Champonnois
' Date      : 25/09/2013
' Purpose   : remove the Test folder contains folders and files
'---------------------------------------------------------------------------------------
Public Sub Test_vtkResetTestFolderContainsFolderAndFiles()

    Dim fso As FileSystemObject, folderTest As Folder
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Test_vtkResetTestFolderContainsFolderAndFiles_Error
    
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "TestFolderReset")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "TestFolderReset" & "\" & "TestSubFolder1")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "TestFolderReset" & "\" & "TestSubFolder2")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "TestFolderReset" & "\" & "TestSubFolder3")
    
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestFolderReset" & "\" & "TestSubFolder1" & "\" & "TestFileSubFolder1.txt")
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestFolderReset" & "\" & "TestSubFolder2" & "\" & "TestFileSubFolder2.txt")
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestFolderReset" & "\" & "TestSubFolder3" & "\" & "TestFileSubFolder3.txt")
    
    resetTestFolder
    
    Set folderTest = fso.GetFolder(VBAToolKit.vtkTestPath)
    
    mAssert.Equals folderTest.Files.count, 0, "All files hasn't been deleted"
    mAssert.Equals folderTest.SubFolders.count, 0, "All folders hasn't been deleted"

    On Error GoTo 0
    Exit Sub

Test_vtkResetTestFolderContainsFolderAndFiles_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkResetTestFolderContainsFolderAndFiles of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkTestPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetNotFileFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetReplaceFileFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplateOpenExcel"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplateExcelNotOpen"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetErrorFolderFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileNotVisibleFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetNotFileFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetReplaceFileFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetNotFolderFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetErrorFileFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetReplaceFolderFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetNotFolderFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetReplaceFolderFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderNotVisibleFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderNotVisibleInFolderFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderCompleteTreeFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkResetTestFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkResetTestFolderFileOpen"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkResetTestFolderEmptyFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkResetTestFolderContainsFiles"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkResetTestFolderContainsFolderAndFiles"
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
        Case "Test_vtkGetFolderCompleteTreeFromTemplate": Test_vtkGetFolderCompleteTreeFromTemplate
        Case "Test_vtkResetTestFolder": Test_vtkResetTestFolder
        Case "Test_vtkResetTestFolderFileOpen": Test_vtkResetTestFolderFileOpen
        Case "Test_vtkResetTestFolderEmptyFolder": Test_vtkResetTestFolderEmptyFolder
        Case "Test_vtkResetTestFolderContainsFiles": Test_vtkResetTestFolderContainsFiles
        Case "Test_vtkResetTestFolderContainsFolderAndFiles": Test_vtkResetTestFolderContainsFolderAndFiles
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

