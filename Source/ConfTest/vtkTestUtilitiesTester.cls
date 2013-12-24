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
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkTestPath of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileFromTemplate_OK
' Author    : Champonnois
' Purpose   : Verify the copy file from template in the folder test
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileFromTemplate_OK()

    Dim wbb As Workbook
    
    On Error GoTo Test_vtkGetFileFromTemplate_OK_Error
    
    Set wbb = getTestFileFromTemplate(fileName:="TestModule.bas")
    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "TestModule.bas"), "TestModule.bas", "The file TestModule should be copied in folder Test"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetFileFromTemplate_OK_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetFileFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileFromTemplate_NoFile
' Author    : Champonnois
' Purpose   : Verifying an error is displayed if the file template doesn't exist
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileFromTemplate_NoFile()

    Dim wbb As Workbook
    
    On Error GoTo Test_vtkGetFileFromTemplate_NoFile_Error
    
    Set wbb = getTestFileFromTemplate(fileName:="Error.bas")
    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "Error.bas"), "", "The inexistant Error.bas template file shouldn't be copied in Test folder"
    On Error GoTo 0
    Exit Sub
Test_vtkGetFileFromTemplate_NoFile_Error:
    If Err.Number <> VTK_FILE_NOT_FOUND Then mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetNotFileFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileFromTemplate_ReplaceFile
' Author    : Champonnois
' Purpose   : Test the replacement of an existing file
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileFromTemplate_ReplaceFile()

    Dim fso As New FileSystemObject

    On Error GoTo Test_vtkGetFileFromTemplate_ReplaceFile_Error
    
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestModule.bas")
    
    getTestFileFromTemplate fileName:="TestModule.bas"
    
    mAssert.Should vtkTextFileReader(VBAToolKit.vtkTestPath & "\" & "TestModule.bas") <> "", "The copied file hasn't replaced the existing empty one"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetFileFromTemplate_ReplaceFile_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetReplaceFileFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileFromTemplate_OpenExcel
' Author    : Champonnois
' Purpose   : After copying an Excel file, verify if the file is opened
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileFromTemplate_OpenExcel()

    Dim wbb As Workbook

    On Error GoTo Test_vtkGetFileFromTemplate_OpenExcel_Error

    Set wbb = getTestFileFromTemplate(fileName:="UninitializedProject.xlsm", openExcel:=True)

    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "UninitializedProject.xlsm"), "UninitializedProject.xlsm", "The file UninitializedProject.xlsm should be copied in folder Test"
    mAssert.Equals wbb.path, VBAToolKit.vtkTestPath, "the Excel file isn't open in the right directory"
    mAssert.Should VtkWorkbookIsOpen(wbb.name), "The file Excel isn't open"

    wbb.Close (False)

    On Error GoTo 0

    Exit Sub

Test_vtkGetFileFromTemplate_OpenExcel_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetFileFromTemplateOpenExcel of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileFromTemplate_ExcelNotOpen
' Author    : Champonnois
' Purpose   : After copying an Excel file, verify if the file isn't opened
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileFromTemplate_ExcelNotOpen()

    Dim wbb As Workbook

    On Error GoTo Test_vtkGetFileFromTemplate_ExcelNotOpen_Error

    Set wbb = getTestFileFromTemplate(fileName:="UninitializedProject.xlsm", openExcel:=False)

    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "UninitializedProject.xlsm"), "UninitializedProject.xlsm", "The file UninitializedProject.xlsm should be copied in folder Test"
    mAssert.Should wbb Is Nothing, "The workbook shouldn't  exist"
    mAssert.Should Not VtkWorkbookIsOpen("UninitializedProject.xlsm"), "the file Excel shouldn't be open"

    On Error GoTo 0

    Exit Sub

Test_vtkGetFileFromTemplate_ExcelNotOpen_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetFileFromTemplateExcelNotOpen of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFileFromTemplate_Folder
' Author    : Champonnois
' Purpose   : Verify that a folder is not copied by the vtkGetFileFromTemplate function
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFileFromTemplate_Folder()
    
    Dim wbb As Workbook

   On Error Resume Next
    Set wbb = getTestFileFromTemplate(fileName:="Folder")
    mAssert.Equals Err.Number, VTK_DOESNT_COPY_FOLDER, "The proper error must be raised when trying to copy a folder"
    
   On Error GoTo Test_vtkGetFileFromTemplate_Folder_Error

    mAssert.Should wbb Is Nothing, "The workbook shouldn't exist"
    mAssert.Equals Dir(VBAToolKit.vtkTestPath & "\" & "Folder"), "", "The folder was copied"
    
   On Error GoTo 0
    Exit Sub

Test_vtkGetFileFromTemplate_Folder_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetErrorFolderFromTemplate of Class Module vtkTestUtilitiesTester"
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
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetFileFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
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
    mAssert.Equals Err.Number, VTK_FILE_NOT_FOUND, "The proper error must be raised when trying to copy a folder"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetNotFileFromTemplate_ParamDestination_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetNotFileFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetReplaceFileFromTemplate_ParamDestination
' Author    : Champonnois
' Purpose   : Same procedure as the replace of a file, but with a new name for the copied file
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetReplaceFileFromTemplate_ParamDestination()
    
    Dim fso As New FileSystemObject, wbb As Workbook, FileLenBeforeCopy As Long, FileLenAfterCopy As Long

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
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetReplaceFileFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next

End Sub

Private Sub assertFolderIsCopied(folderName As String, mAssert As IAssert)
    Dim fso As New FileSystemObject
    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & folderName), "The folder " & folderName & " should have been copied in folder Test"
    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & folderName & "\SubFolder"), "The folder " & folderName & "\SubFolder should have been copied in folder Test"
    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & folderName & "\TestClass.cls"), "The file " & folderName & "\TestClass.cls should have been copied in folder Test"
    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & folderName & "\SubFolder\TestModule.bas"), "The file " & folderName & "\SubFolder\TestModule.bas should have been copied in folder Test"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderFromTemplate
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verify the copy folder from template in the folder test
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderFromTemplate()
    
    On Error GoTo Test_vtkGetFolderFromTemplate_Error
    
    getTestFolderFromTemplate folderName:="Folder"
    
    assertFolderIsCopied folderName:="Folder", mAssert:=mAssert
'    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Folder"), "The folder Folder should have been copied in folder Test"
'    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Folder\SubFolder"), "The folder Folder\SubFolder should have been copied in folder Test"
'    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & "Folder\TestClass.cls"), "The file Folder\TestClass.cls should have been copied in folder Test"
'    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & "Folder\SubFolder\TestModule.bas"), "The file Folder\SubFolder\TestModule.bas should have been copied in folder Test"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetFolderFromTemplate_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetFolderFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderFromTemplate_NoFolder
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verifying an error is displayed if one doesn't possess the folder in template
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderFromTemplate_NoFolder()
    Dim fso As New FileSystemObject
    
   On Error Resume Next
    getTestFolderFromTemplate folderName:="Error"
    mAssert.Equals Err.Number, VTK_FOLDER_NOT_FOUND, "The proper error must be raised when trying to copy a folder"
    
   On Error GoTo Test_vtkGetFolderFromTemplate_NoFolder_Error
    mAssert.Equals fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Error"), False, "The folder Error shouldn't have to be copied into folder Test"
    
    On Error GoTo 0
    
    Exit Sub
Test_vtkGetFolderFromTemplate_NoFolder_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetNotFolderFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderFromTemplate_CopyFile
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Verifying that you can't copy the file and displays an error
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderFromTemplate_CopyFile()
    Dim fso As New FileSystemObject
    
   On Error Resume Next
    getTestFolderFromTemplate folderName:="Module.bas"
    mAssert.Equals Err.Number, VTK_FOLDER_NOT_FOUND, "The proper error must be raised when trying to copy a file instead of a folder"

   On Error GoTo Test_vtkGetFolderFromTemplate_CopyFile_Error
    mAssert.Equals fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Module.bas"), False, "The file Module.bas shouldn't have to be copied into folder Test"
    
    On Error GoTo 0
    Exit Sub
Test_vtkGetFolderFromTemplate_CopyFile_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetErrorFolderFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderFromTemplate_ReplaceFolder
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Test the replacement of an existing folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderFromTemplate_ReplaceFolder()
    Dim fso As New FileSystemObject, FolderBeforeCopy As Folder, FolderAfterCopy As Folder, SizeFolderBeforeCopy As Integer, SizeFolderAfterCopy As Integer

    On Error GoTo Test_vtkGetFolderFromTemplate_ReplaceFolder_Error
    
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "Folder")
    
    Set FolderBeforeCopy = fso.GetFolder(VBAToolKit.vtkTestPath & "\" & "Folder")
    SizeFolderBeforeCopy = FolderBeforeCopy.Size
    
    getTestFolderFromTemplate folderName:="Folder"
    
    Set FolderAfterCopy = fso.GetFolder(VBAToolKit.vtkTestPath & "\" & "Folder")
    SizeFolderAfterCopy = FolderAfterCopy.Size
    
    mAssert.Should SizeFolderBeforeCopy < SizeFolderAfterCopy, "the copied file doesn't replace the existing file"
    assertFolderIsCopied folderName:="Folder", mAssert:=mAssert
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetFolderFromTemplate_ReplaceFolder_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetReplaceFolderFromTemplate of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderFromTemplate_ParamDestination
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Same procedure as the copy of a folder, but with a new name for the copied folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderFromTemplate_ParamDestination()
    Dim fso As New FileSystemObject
    
    On Error GoTo Test_vtkGetFolderFromTemplate_ParamDestination_Error
    
    getTestFolderFromTemplate folderName:="Folder", destinationName:="TestRenameFolder"
    
    assertFolderIsCopied folderName:="TestRenameFolder", mAssert:=mAssert
    mAssert.Equals fso.folderExists(VBAToolKit.vtkTestPath & "\" & "Folder"), False, "The folder Folder shouldn't have to be copied into folder Test with the same name"
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetFolderFromTemplate_ParamDestination_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetFolderFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderFromTemplate_NoFolderParamDestination
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Same procedure as the copy of a nonexistent folder, but with a new name for the copied folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderFromTemplate_NoFolderParamDestination()
    Dim fso As New FileSystemObject
    
   On Error Resume Next
    getTestFolderFromTemplate folderName:="Error", destinationName:="TestRenameError"
    mAssert.Equals Err.Number, VTK_FOLDER_NOT_FOUND, "The proper error must be raised when trying to copy a folder"
    
   On Error GoTo Test_vtkGetFolderFromTemplate_NoFolderParamDestination_Error
    mAssert.Equals fso.folderExists(VBAToolKit.vtkTestPath & "\" & "TestRenameError"), False, "Bad test the folder Folder hasn't been copied in folder Test"
    
    On Error GoTo 0
    
    Exit Sub
Test_vtkGetFolderFromTemplate_NoFolderParamDestination_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetNotFolderFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkGetFolderFromTemplate_ReplaceFolderParamDestination
' Author    : Champonnois
' Date      : 23/09/2013
' Purpose   : Same procedure as the replace of a folder, but with a new name for the copied folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkGetFolderFromTemplate_ReplaceFolderParamDestination()
    Dim fso As New FileSystemObject, FolderBeforeCopy As Folder, FolderAfterCopy As Folder, SizeFolderBeforeCopy As Integer, SizeFolderAfterCopy As Integer

   On Error GoTo Test_vtkGetFolderFromTemplate_ReplaceFolderParamDestination_Error
    
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "TestRenameFolder")
    
    Set FolderBeforeCopy = fso.GetFolder(VBAToolKit.vtkTestPath & "\" & "TestRenameFolder")
    SizeFolderBeforeCopy = FolderBeforeCopy.Size
    
    getTestFolderFromTemplate folderName:="Folder", destinationName:="TestRenameFolder"
    
    Set FolderAfterCopy = fso.GetFolder(VBAToolKit.vtkTestPath & "\" & "TestRenameFolder")
    SizeFolderAfterCopy = FolderAfterCopy.Size
    
    mAssert.Should SizeFolderBeforeCopy < SizeFolderAfterCopy, "the copied file doesn't replace the existing file"
    assertFolderIsCopied folderName:="TestRenameFolder", mAssert:=mAssert
    
    On Error GoTo 0
    Exit Sub
    
Test_vtkGetFolderFromTemplate_ReplaceFolderParamDestination_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGetReplaceFolderFromTemplate_ParamDestination of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

Private Sub assertTestFolderIsEmpty(mAssert As IAssert)
    Dim fso As New FileSystemObject, folderTest As Folder
    Set folderTest = fso.GetFolder(VBAToolKit.vtkTestPath)
    
    mAssert.Equals folderTest.Files.count, 0, "All files hasn't been deleted"
    mAssert.Equals folderTest.SubFolders.count, 0, "All folders hasn't been deleted"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Test_vtkResetTestFolder_OneFile
' Author    : Champonnois
' Date      : 25/09/2013
' Purpose   : remove the file in Test folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkResetTestFolder_OneFile()
    Dim fso As New FileSystemObject
    
   On Error GoTo Test_vtkResetTestFolder_OneFile_Error

    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "TestResetFile.bas")
    
    resetTestFolder
    
    assertTestFolderIsEmpty mAssert:=mAssert

    On Error GoTo 0
    Exit Sub

Test_vtkResetTestFolder_OneFile_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkResetTestFolder of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkResetTestFolder_FileOpen
' Author    : Champonnois
' Date      : 25/09/2013
' Purpose   : remove the file that is open in Test folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkResetTestFolder_FileOpen()
    Dim fso As New FileSystemObject, wbb As Workbook
    
   On Error GoTo Test_vtkResetTestFolder_FileOpen_Error
    Set wbb = getTestFileFromTemplate(fileName:="UninitializedProject.xlsm", openExcel:=True)

   On Error Resume Next
    resetTestFolder
    mAssert.Equals Err.Number, VTK_FILE_OPEN_OR_LOCKED, "The proper error must be raised when attempting to clean folder when a file is opened"
    
   On Error GoTo Test_vtkResetTestFolder_FileOpen_Error
    mAssert.Should VtkWorkbookIsOpen("UninitializedProject.xlsm"), "The workbook UninitializedProject.xlsm must not be closed when attempting to clean the Test folder"

   On Error GoTo 0
    vtkCloseAndKillWorkbook wbb
    Exit Sub

Test_vtkResetTestFolder_FileOpen_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkResetTestFolderFileOpen of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkResetTestFolder_EmptyFolder
' Author    : Champonnois
' Date      : 25/09/2013
' Purpose   : remove the empty folder in Test folder
'---------------------------------------------------------------------------------------
Public Sub Test_vtkResetTestFolder_EmptyFolder()
    
   On Error GoTo Test_vtkResetTestFolder_EmptyFolder_Error
    
    resetTestFolder
    
    assertTestFolderIsEmpty mAssert:=mAssert
    
    On Error GoTo 0
    Exit Sub

Test_vtkResetTestFolder_EmptyFolder_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkResetTestFolderEmptyFolder of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkResetTestFolder_PopulatedFolder
' Author    : Champonnois
' Date      : 25/09/2013
' Purpose   : remove the Test folder contains files
'---------------------------------------------------------------------------------------
Public Sub Test_vtkResetTestFolder_PopulatedFolder()

   On Error GoTo Test_vtkResetTestFolder_PopulatedFolder_Error
    
    getTestFolderFromTemplate folderName:="Folder"
    
    resetTestFolder
    
    assertTestFolderIsEmpty mAssert:=mAssert

    On Error GoTo 0
    Exit Sub

Test_vtkResetTestFolder_PopulatedFolder_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkResetTestFolderContainsFiles of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkResetTestFolder_PopulatedFoldersAndFiles
' Author    : Imbert
' Date      : 24/12/2013
' Purpose   : remove the Test folder contains files and populated folders
'---------------------------------------------------------------------------------------
Public Sub Test_vtkResetTestFolder_PopulatedFoldersAndFiles()
    Dim fso As New FileSystemObject, fileObject As File

   On Error GoTo Test_vtkResetTestFolder_PopulatedFoldersAndFiles_Error
    
    getTestFolderFromTemplate folderName:="Folder", destinationName:="FirstFolder"
    getTestFolderFromTemplate folderName:="Folder", destinationName:="SecondFolder"
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "FirstFile.bat")
    Set fileObject = fso.GetFile(VBAToolKit.vtkTestPath & "\" & "FirstFile.bat")
    fileObject.Attributes = fileObject.Attributes Or 1  '   Set The Read Only attribute
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "SecondFile.bas")

    resetTestFolder
    
    assertTestFolderIsEmpty mAssert:=mAssert

    On Error GoTo 0
    Exit Sub

Test_vtkResetTestFolder_PopulatedFoldersAndFiles_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkResetTestFolderContainsFiles of Class Module vtkTestUtilitiesTester"
    Resume Next
    
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkTestPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate_OK"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate_NoFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate_ReplaceFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate_OpenExcel"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate_ExcelNotOpen"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate_Folder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFileFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetNotFileFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetReplaceFileFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate_NoFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate_CopyFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate_ReplaceFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate_ParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate_NoFolderParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGetFolderFromTemplate_ReplaceFolderParamDestination"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkResetTestFolder_OneFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkResetTestFolder_FileOpen"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkResetTestFolder_EmptyFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkResetTestFolder_PopulatedFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkResetTestFolder_PopulatedFoldersAndFiles"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkTestPath": Test_vtkTestPath
        Case "Test_vtkGetFileFromTemplate_OK": Test_vtkGetFileFromTemplate_OK
        Case "Test_vtkGetFileFromTemplate_NoFile": Test_vtkGetFileFromTemplate_NoFile
        Case "Test_vtkGetFileFromTemplate_ReplaceFile": Test_vtkGetFileFromTemplate_ReplaceFile
        Case "Test_vtkGetFileFromTemplate_OpenExcel": Test_vtkGetFileFromTemplate_OpenExcel
        Case "Test_vtkGetFileFromTemplate_ExcelNotOpen": Test_vtkGetFileFromTemplate_ExcelNotOpen
        Case "Test_vtkGetFileFromTemplate_Folder": Test_vtkGetFileFromTemplate_Folder
        Case "Test_vtkGetFileFromTemplate_ParamDestination": Test_vtkGetFileFromTemplate_ParamDestination
        Case "Test_vtkGetNotFileFromTemplate_ParamDestination": Test_vtkGetNotFileFromTemplate_ParamDestination
        Case "Test_vtkGetReplaceFileFromTemplate_ParamDestination": Test_vtkGetReplaceFileFromTemplate_ParamDestination
        Case "Test_vtkGetFolderFromTemplate": Test_vtkGetFolderFromTemplate
        Case "Test_vtkGetFolderFromTemplate_NoFolder": Test_vtkGetFolderFromTemplate_NoFolder
        Case "Test_vtkGetFolderFromTemplate_CopyFile": Test_vtkGetFolderFromTemplate_CopyFile
        Case "Test_vtkGetFolderFromTemplate_ReplaceFolder": Test_vtkGetFolderFromTemplate_ReplaceFolder
        Case "Test_vtkGetFolderFromTemplate_ParamDestination": Test_vtkGetFolderFromTemplate_ParamDestination
        Case "Test_vtkGetFolderFromTemplate_NoFolderParamDestination": Test_vtkGetFolderFromTemplate_NoFolderParamDestination
        Case "Test_vtkGetFolderFromTemplate_ReplaceFolderParamDestination": Test_vtkGetFolderFromTemplate_ReplaceFolderParamDestination
        Case "Test_vtkResetTestFolder_OneFile": Test_vtkResetTestFolder_OneFile
        Case "Test_vtkResetTestFolder_FileOpen": Test_vtkResetTestFolder_FileOpen
        Case "Test_vtkResetTestFolder_EmptyFolder": Test_vtkResetTestFolder_EmptyFolder
        Case "Test_vtkResetTestFolder_PopulatedFolder": Test_vtkResetTestFolder_PopulatedFolder
        Case "Test_vtkResetTestFolder_PopulatedFoldersAndFiles": Test_vtkResetTestFolder_PopulatedFoldersAndFiles
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

