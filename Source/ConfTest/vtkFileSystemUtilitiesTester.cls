VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkFileSystemUtilitiesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    fso.CreateFolder (vtkTestPath & "\" & "visible_folder")
    fso.CreateFolder (vtkTestPath & "\" & "visible_folder" & "\" & "visible_folder")
    fso.CreateFolder (vtkTestPath & "\" & "hidden_folder")
    fso.CreateFolder (vtkTestPath & "\" & "hidden_folder" & "\" & "visible_folder")
    fso.CreateFolder (vtkTestPath & "\" & "testFolder0")
    fso.CreateFolder (vtkTestPath & "\" & "testFolder0" & "\" & "testFolder11")
    fso.CreateFolder (vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1")
    fso.CreateFolder (vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2")
    fso.CreateFolder (vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2" & "\" & "testFolder3")
    ' We create three text files to try different configurations
    '   - testFolder0  : text file and 2 subfolders
    '   - testFolder1  : text file and one subfolder
    '   - testFolder11 : empty
    '   - testFolder2  : subfolder only
    '   - testFolder3  : text file only
    fso.CreateTextFile (vtkTestPath & "\" & "testFolder0" & "\" & "textFile0.txt")
    fso.CreateTextFile (vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "textFile1.txt")
    fso.CreateTextFile (vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2" & _
    "\" & "testFolder3" & "\" & "textFile3.txt")
    
    '------------------------------------------------------------------------------
    
    'Creating other files and folders for simple tests
    fso.CreateTextFile (vtkTestPath & "\" & "testFolder0" & "\" & "hidden.txt")
    fso.CreateTextFile (vtkTestPath & "\" & "testFolder0" & "\" & "not_hidden.txt")
    fso.CreateFolder (vtkTestPath & "\" & "testFolder0" & "\" & "hidden_folder")
    SetAttr vtkTestPath & "\" & "testFolder0" & "\" & "hidden.txt", vbHidden
    SetAttr vtkTestPath & "\" & "testFolder0" & "\" & "hidden_folder", vbHidden
End Sub

Private Sub ITestCase_TearDown()
    
    On Error Resume Next
    
    fso = Nothing

    vtkCleanFolder (vtkTestPath)
    
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
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


Public Sub Test_vtkCleanFolder_correctPathAreFoldersDeleted()
    Dim fso As New Scripting.FileSystemObject
    Dim testFolder As Scripting.Folder

    Set testFolder = fso.GetFolder(vtkTestPath & "\" & "testFolder0")

    vtkCleanFolder (testFolder.path)
    
    ' (testFolder.SubFolders.Count = 0) And (testFolder.Files.Count = 0) does look at the hidden files.
    mAssert.Should (testFolder.SubFolders.Count = 0) And (testFolder.Files.Count = 0) = True, "All files and folders must be deleted"
End Sub
'
Public Sub Test_vtkCleanFolder_correctPathReturnValue()
    Dim testFolder As String
    testFolder = vtkTestPath & "\" & "testFolder0"
    
    Dim retval As Integer
    retval = vtkCleanFolder(testFolder)

    mAssert.Equals retval, VTK_OK, "Wrong return value"
End Sub

Public Sub Test_vtkCleanFolder_wrongPath()
    Dim testFolder As String
    testFolder = vtkTestPath & "\" & "no_folder"
    
    Dim retval As Integer
    retval = vtkCleanFolder(testFolder)

    mAssert.Equals retval, 76, "Return value must be 76 : path not found"
End Sub

Public Sub Test_vtkCleanFolder_file()
    Dim filePath As String
    filePath = vtkTestPath & "\" & "hidden.txt"
    
    Dim retval As Integer
    retval = vtkCleanFolder(filePath)

    mAssert.Equals retval, 76, "Return value must be 76 : path not found"
End Sub


Public Sub Test_vtkDeleteFolder_correctPathAreFoldersDeleted()
    Dim testFolder As String
    testFolder = vtkTestPath & "\" & "testFolder0"
    
    vtkDeleteFolder testFolder
    
    mAssert.Equals vtkDoesFolderExist(testFolder), False, "Folder must be deleted"
End Sub
'
Public Sub Test_vtkDeleteFolder_correctPathReturnValue()
    Dim testFolder As String
    testFolder = vtkTestPath & "\" & "testFolder0"
    
    Dim retval As Integer
    retval = vtkDeleteFolder(testFolder)

    mAssert.Equals retval, VTK_OK, "Wrong return value"
End Sub

Public Sub Test_vtkDeleteFolder_wrongPath()

    Dim testFolder As String
    testFolder = vtkTestPath & "\" & "no_folder"
    
    Dim retval As Integer
    retval = vtkDeleteFolder(testFolder)

    mAssert.Equals retval, 76, "Return value must be 76 : path not found"
End Sub

Public Sub Test_vtkDeleteFolder_file()
    Dim filePath As String
    filePath = vtkTestPath & "\" & "hidden.txt"
    
    Dim retval As Integer
    retval = vtkCleanFolder(filePath)

    mAssert.Equals retval, 76, "Return value must be 76 : path not found"
End Sub


Public Sub Test_vtkDoesFolderExist_visibleFolder()
    mAssert.Should vtkDoesFolderExist(vtkTestPath & "\" & "testFolder0") = True, "Folder exists."
End Sub

Public Sub Test_vtkDoesFolderExist_hiddenFolder()
    mAssert.Should vtkDoesFolderExist(vtkTestPath & "\" & "hidden_folder") = True, "Folder exists."
End Sub

Public Sub Test_vtkDoesFolderExist_wrongPath()
    mAssert.Should vtkDoesFolderExist(vtkTestPath & "\" & "no_folder") = False, "Folder does not exist."
End Sub

Public Sub Test_vtkDoesFolderExist_visibleFile()
    mAssert.Should vtkDoesFolderExist(vtkTestPath & "\" & "not_hidden.txt") = False, "Is not a folder."
End Sub

Public Sub Test_vtkDoesFolderExist_hiddenFile()
    mAssert.Should vtkDoesFolderExist(vtkTestPath & "\" & "hidden.txt") = False, "Is not a folder."
End Sub




