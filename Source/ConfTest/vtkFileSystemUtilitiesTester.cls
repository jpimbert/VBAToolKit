VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkFileSystemUtilitiesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkFileSystemUtilitiesTester
' Author    : Lucas Vitorino
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
Private fso As FileSystemObject

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
    
    Set fso = New FileSystemObject
    
    'Creating a hierarchy of folders (mainly for vtkCleanFolder)
    ' -----------------------------------------------------------------------------
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "visible_folder")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "visible_folder" & "\" & "visible_folder")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "hidden_folder")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "hidden_folder" & "\" & "visible_folder")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "testFolder0")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "testFolder11")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2" & "\" & "testFolder3")
    ' We create three text files to try different configurations
    '   - testFolder0  : text file and 2 subfolders
    '   - testFolder1  : text file and one subfolder
    '   - testFolder11 : empty
    '   - testFolder2  : subfolder only
    '   - testFolder3  : text file only
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "textFile0.txt")
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "textFile1.txt")
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2" & _
    "\" & "testFolder3" & "\" & "textFile3.txt")
    
    '------------------------------------------------------------------------------
    
    'Creating other files and folders for simple tests
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "hidden.txt")
    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "not_hidden.txt")
    fso.CreateFolder (VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "hidden_folder")
    SetAttr VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "hidden.txt", vbHidden
    SetAttr VBAToolKit.vtkTestPath & "\" & "testFolder0" & "\" & "hidden_folder", vbHidden
End Sub

Private Sub ITestCase_TearDown()
    
    On Error Resume Next
    
    fso = Nothing

    VBAToolKit.resetTestFolder
    
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCleanFolder_correctPathAreFoldersDeleted"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCleanFolder_correctPathReturnValue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCleanFolder_wrongPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCleanFolder_file"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDeleteFolder_correctPathAreFoldersDeleted"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDeleteFolder_correctPathReturnValue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDeleteFolder_wrongPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDeleteFolder_file"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDoesFolderExist_visibleFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDoesFolderExist_hiddenFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDoesFolderExist_wrongPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDoesFolderExist_visibleFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDoesFolderExist_hiddenFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateFolderPath_folderName1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateFolderPath_folderName2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateFolderPath_fileName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateFolderPath_relativePath1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateFolderPath_relativePath2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateFolderPath_wrongDrive"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateFolderPath_noDotsInPath_dot"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateFolderPath_noDotsInPath_doubleDot"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateFolderPath_noDotsInPath_tripleDot"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkCleanFolder_correctPathAreFoldersDeleted": Test_vtkCleanFolder_correctPathAreFoldersDeleted
        Case "Test_vtkCleanFolder_correctPathReturnValue": Test_vtkCleanFolder_correctPathReturnValue
        Case "Test_vtkCleanFolder_wrongPath": Test_vtkCleanFolder_wrongPath
        Case "Test_vtkCleanFolder_file": Test_vtkCleanFolder_file
        Case "Test_vtkDeleteFolder_correctPathAreFoldersDeleted": Test_vtkDeleteFolder_correctPathAreFoldersDeleted
        Case "Test_vtkDeleteFolder_correctPathReturnValue": Test_vtkDeleteFolder_correctPathReturnValue
        Case "Test_vtkDeleteFolder_wrongPath": Test_vtkDeleteFolder_wrongPath
        Case "Test_vtkDeleteFolder_file": Test_vtkDeleteFolder_file
        Case "Test_vtkDoesFolderExist_visibleFolder": Test_vtkDoesFolderExist_visibleFolder
        Case "Test_vtkDoesFolderExist_hiddenFolder": Test_vtkDoesFolderExist_hiddenFolder
        Case "Test_vtkDoesFolderExist_wrongPath": Test_vtkDoesFolderExist_wrongPath
        Case "Test_vtkDoesFolderExist_visibleFile": Test_vtkDoesFolderExist_visibleFile
        Case "Test_vtkDoesFolderExist_hiddenFile": Test_vtkDoesFolderExist_hiddenFile
        Case "Test_vtkCreateFolderPath_folderName1": Test_vtkCreateFolderPath_folderName1
        Case "Test_vtkCreateFolderPath_folderName2": Test_vtkCreateFolderPath_folderName2
        Case "Test_vtkCreateFolderPath_fileName": Test_vtkCreateFolderPath_fileName
        Case "Test_vtkCreateFolderPath_relativePath1": Test_vtkCreateFolderPath_relativePath1
        Case "Test_vtkCreateFolderPath_relativePath2": Test_vtkCreateFolderPath_relativePath2
        Case "Test_vtkCreateFolderPath_wrongDrive": Test_vtkCreateFolderPath_wrongDrive
        Case "Test_vtkCreateFolderPath_noDotsInPath_dot": Test_vtkCreateFolderPath_noDotsInPath_dot
        Case "Test_vtkCreateFolderPath_noDotsInPath_doubleDot": Test_vtkCreateFolderPath_noDotsInPath_doubleDot
        Case "Test_vtkCreateFolderPath_noDotsInPath_tripleDot": Test_vtkCreateFolderPath_noDotsInPath_tripleDot
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


Public Sub Test_vtkCleanFolder_correctPathAreFoldersDeleted()
    Dim fso As New Scripting.FileSystemObject
    Dim testFolder As Scripting.Folder

    Set testFolder = fso.GetFolder(VBAToolKit.vtkTestPath)

    vtkCleanFolder (testFolder.path)
    
    ' (testFolder.SubFolders.Count = 0) And (testFolder.Files.Count = 0) does look at the hidden files.
    mAssert.Should (testFolder.SubFolders.Count = 0) And (testFolder.Files.Count = 0), "All files and folders must be deleted"
End Sub
'
Public Sub Test_vtkCleanFolder_correctPathReturnValue()
    Dim testFolder As String
    testFolder = VBAToolKit.vtkTestPath & "\" & "testFolder0"
    
    Dim retVal As Integer
    retVal = vtkCleanFolder(testFolder)

    mAssert.Equals retVal, VTK_OK, "Wrong return value"
End Sub

Public Sub Test_vtkCleanFolder_wrongPath()
    Dim testFolder As String
    testFolder = VBAToolKit.vtkTestPath & "\" & "no_folder"
    
    Dim retVal As Integer
    retVal = vtkCleanFolder(testFolder)

    mAssert.Equals retVal, 76, "Return value must be 76 : path not found"
End Sub

Public Sub Test_vtkCleanFolder_file()
    Dim filePath As String
    filePath = VBAToolKit.vtkTestPath & "\" & "hidden.txt"
    
    Dim retVal As Integer
    retVal = vtkCleanFolder(filePath)

    mAssert.Equals retVal, 76, "Return value must be 76 : path not found"
End Sub


Public Sub Test_vtkDeleteFolder_correctPathAreFoldersDeleted()
    Dim testFolder As String
    testFolder = VBAToolKit.vtkTestPath & "\" & "testFolder0"
    
    vtkDeleteFolder testFolder
    
    mAssert.Equals vtkDoesFolderExist(testFolder), False, "Folder must be deleted"
End Sub
'
Public Sub Test_vtkDeleteFolder_correctPathReturnValue()
    Dim testFolder As String
    testFolder = VBAToolKit.vtkTestPath & "\" & "testFolder0"
    
    Dim retVal As Integer
    retVal = vtkDeleteFolder(testFolder)

    mAssert.Equals retVal, VTK_OK, "Wrong return value"
End Sub

Public Sub Test_vtkDeleteFolder_wrongPath()

    Dim testFolder As String
    testFolder = VBAToolKit.vtkTestPath & "\" & "no_folder"
    
    Dim retVal As Integer
    retVal = vtkDeleteFolder(testFolder)

    mAssert.Equals retVal, 76, "Return value must be 76 : path not found"
End Sub

Public Sub Test_vtkDeleteFolder_file()
    Dim filePath As String
    filePath = VBAToolKit.vtkTestPath & "\" & "hidden.txt"
    
    Dim retVal As Integer
    retVal = vtkCleanFolder(filePath)

    mAssert.Equals retVal, 76, "Return value must be 76 : path not found"
End Sub


Public Sub Test_vtkDoesFolderExist_visibleFolder()
    mAssert.Should vtkDoesFolderExist(VBAToolKit.vtkTestPath & "\" & "testFolder0") = True, "Folder exists."
End Sub

Public Sub Test_vtkDoesFolderExist_hiddenFolder()
    mAssert.Should vtkDoesFolderExist(VBAToolKit.vtkTestPath & "\" & "hidden_folder") = True, "Folder exists."
End Sub

Public Sub Test_vtkDoesFolderExist_wrongPath()
    mAssert.Should vtkDoesFolderExist(VBAToolKit.vtkTestPath & "\" & "no_folder") = False, "Folder does not exist."
End Sub

Public Sub Test_vtkDoesFolderExist_visibleFile()
    mAssert.Should vtkDoesFolderExist(VBAToolKit.vtkTestPath & "\" & "not_hidden.txt") = False, "Is not a folder."
End Sub

Public Sub Test_vtkDoesFolderExist_hiddenFile()
    mAssert.Should vtkDoesFolderExist(VBAToolKit.vtkTestPath & "\" & "hidden.txt") = False, "Is not a folder."
End Sub



Public Sub Test_vtkCreateFolderPath_folderName1()
' correct folder path : folder should be created

    On Error GoTo Test_vtkCreateFolderPath_folderName_Error

    Dim folderPath As String
    folderPath = fso.BuildPath(VBAToolKit.vtkTestPath, "testFolder10")

    vtkCreateFolderPath folderPath

    mAssert.Should fso.folderExists(folderPath), "Folder should have been created."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateFolderPath_folderName_Error:
    Err.Source = "Test_vtkCreateFolderPath_folderName of module vtkFileSystemUtilitiesTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkCreateFolderPath_folderName2()
' correct folder path : folder tree should be created

    On Error GoTo Test_vtkCreateFolderPath_folderName_Error

    Dim folderPath As String
    folderPath = fso.BuildPath(VBAToolKit.vtkTestPath, "testFolder10\testFolder20")

    vtkCreateFolderPath folderPath

    mAssert.Should fso.folderExists(folderPath), "Folder should have been created."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateFolderPath_folderName_Error:
    Err.Source = "Test_vtkCreateFolderPath_folderName of module vtkFileSystemUtilitiesTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkCreateFolderPath_fileName()
' correct file path : folder tree should be created, but not the file

    On Error GoTo Test_vtkCreateFolderPath_fileName_Error

    Dim filePath As String
    Dim folderPath As String
    filePath = fso.BuildPath(VBAToolKit.vtkTestPath, "testFolder10\testFolder20\dummy.txt")
    folderPath = fso.GetParentFolderName(filePath)

    vtkCreateFolderPath filePath

    mAssert.Should fso.folderExists(folderPath), _
                   "Folder tree should have been created."

    mAssert.Should Not fso.folderExists(filePath), _
                   "No folder with the same name as the file should have been created"

    mAssert.Should Not fso.FileExists(filePath), _
                   "File should not have been created : function is just supposed to create the folder tree."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateFolderPath_fileName_Error:
    Err.Source = "Test_vtkCreateFolderPath_fileName of module vtkFileSystemUtilitiesTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkCreateFolderPath_relativePath1()
' relative folder path : VTK_FORBIDDEN_PARAMETER should be raised

    On Error GoTo Test_vtkCreateFolderPath_relativePath1_Error

    Dim error As Long: error = 0

    Dim folderPath As String
    folderPath = "testFolder10"

    vtkCreateFolderPath folderPath

    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER, _
                   "Function should not accept relative paths"

    On Error GoTo 0
    Exit Sub

Test_vtkCreateFolderPath_relativePath1_Error:
    error = Err.Number
    Exit Sub
End Sub


Public Sub Test_vtkCreateFolderPath_relativePath2()
' relative folder path : VTK_FORBIDDEN_PARAMETER should be raised

    On Error GoTo Test_vtkCreateFolderPath_relativePath2_Error

    Dim error As Long: error = 0

    Dim folderPath As String
    folderPath = "\testFolder10"

    vtkCreateFolderPath folderPath

    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER, _
                   "Function should not accept relative paths"

    On Error GoTo 0
    Exit Sub

Test_vtkCreateFolderPath_relativePath2_Error:
    error = Err.Number
    Exit Sub
End Sub



Public Sub Test_vtkCreateFolderPath_wrongDrive()
' absolute path is given on a unmounted drive : VTK_WRONG_FOLDER_PATH is raised

    On Error GoTo Test_vtkCreateFolderPath_wrongDrive_Error

    Dim error As Long: error = 0

    Dim folderPath As String
    folderPath = "O:\testFolder1\testFolder2"

    vtkCreateFolderPath folderPath

    mAssert.Equals error, VTK_WRONG_FOLDER_PATH, _
                   "The drive is not mounted, the call is supposed to fail."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateFolderPath_wrongDrive_Error:
    error = Err.Number
    Exit Sub
End Sub



Public Sub Test_vtkCreateFolderPath_noDotsInPath_dot()
' path with a dot inside is given : VTK_FORBIDDEN_PARAMETER should be raised

    On Error GoTo Test_vtkCreateFolderPath_noDotsInPath_dot_Error

    Dim error As Long: error = 0

    Dim folderPath As String
    folderPath = fso.BuildPath(VBAToolKit.vtkTestPath, ".\testFolder10")

    vtkCreateFolderPath folderPath

    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER, _
                   "The function should not accept dots in its path."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateFolderPath_noDotsInPath_dot_Error:
    error = Err.Number
    Exit Sub
End Sub



Public Sub Test_vtkCreateFolderPath_noDotsInPath_doubleDot()
' path with two dots inside is given : VTK_FORBIDDEN_PARAMETER should be raised

    On Error GoTo Test_vtkCreateFolderPath_noDotsInPath_doubleDot_Error

    Dim error As Long: error = 0

    Dim folderPath As String
    folderPath = fso.BuildPath(VBAToolKit.vtkTestPath, "testFolder10\..\testFolder20")

    vtkCreateFolderPath folderPath

    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER, _
                   "The function should not accept dots in its path."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateFolderPath_noDotsInPath_doubleDot_Error:
    error = Err.Number
    Exit Sub
End Sub


Public Sub Test_vtkCreateFolderPath_noDotsInPath_tripleDot()
' path with three dots inside is given : VTK_FORBIDDEN_PARAMETER should be raised

    On Error GoTo Test_vtkCreateFolderPath_noDotsInPath_tripleDot_Error

    Dim error As Long: error = 0

    Dim folderPath As String
    folderPath = fso.BuildPath(VBAToolKit.vtkTestPath, "testFolder10\...\testFolder20")

    vtkCreateFolderPath folderPath

    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER, _
                   "The function should not accept dots in its path."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateFolderPath_noDotsInPath_tripleDot_Error:
    error = Err.Number
    Exit Sub
End Sub

