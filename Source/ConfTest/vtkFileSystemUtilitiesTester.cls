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
    
    On Error Resume Next
    
    'Creating a hierarchy of folders.
    fso.CreateFolder (vtkTestPath & "\" & "testFolder0")
    fso.CreateFolder (vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1")
    fso.CreateFolder (vtkTestPath & "\" & "testFolder0" & "\" & "testFolder11")
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
    
End Sub

Private Sub ITestCase_TearDown()
    
    On Error Resume Next
    
    fso = Nothing

    ' Kill the files
    Kill vtkTestPath & "\" & "testFolder0" & "\" & "*"
    Kill vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "*"
    Kill vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2" & _
    "\" & "*"
    Kill vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2" & _
    "\" & "testFolder3" & "\" & "*"

    ' Delete the now empty folders
    RmDir vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2" & "\" & "testFolder3"
    RmDir vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2"
    RmDir vtkTestPath & "\" & "testFolder0" & "\" & "testFolder11"
    RmDir vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1"
    RmDir vtkTestPath & "\" & "testFolder0"

End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestvtkCreateFileInDirectory"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCleanFolder__correct_path"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkIsFolderEmpty__empty_folder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkIsFolderEmpty__folder_with_a_subfolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkIsFolderEmpty__folder_with_a_file"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkIsFolderEmpty__folder_with_a_file_and_a_subfolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDoesFolderExist__folder_existing"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDoesFolderExist__folder_not_existing"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestvtkCreateFileInDirectory": TestvtkCreateFileInDirectory
        Case "Test_vtkCleanFolder__correct_path": Test_vtkCleanFolder__correct_path
        Case "Test_vtkIsFolderEmpty__empty_folder": Test_vtkIsFolderEmpty__empty_folder
        Case "Test_vtkIsFolderEmpty__folder_with_a_subfolder": Test_vtkIsFolderEmpty__folder_with_a_subfolder
        Case "Test_vtkIsFolderEmpty__folder_with_a_file": Test_vtkIsFolderEmpty__folder_with_a_file
        Case "Test_vtkIsFolderEmpty__folder_with_a_file_and_a_subfolder": Test_vtkIsFolderEmpty__folder_with_a_file_and_a_subfolder
        Case "Test_vtkDoesFolderExist__folder_existing": Test_vtkDoesFolderExist__folder_existing
        Case "Test_vtkDoesFolderExist__folder_not_existing": Test_vtkDoesFolderExist__folder_not_existing
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestvtkCreateFileInDirectoryFunction
' Author    : Abdelfattah Lahbib
' Purpose   : Test the vtkCreateFileInDirectory function
'               - After calling this function with a functionnal path, the file must be created.
'---------------------------------------------------------------------------------------
'
Public Sub TestvtkCreateFileInDirectory()

Debug.Print "Entering vtkCreateFileInDirectory"
    Dim fso As New FileSystemObject
    vtkCreateFileInDirectory "dummyFile.txt", vtkTestPath & "\" & "\testFolder0"
    mAssert.Should fso.FileExists(vtkTestPath & "\" & "\testFolder0" & "\" & "dummyFile.txt") = True, "File must be created"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_vtkCleanFolder_correct_path
' Author    : Lucas Vitorino
' Purpose   : Test the vtkCleanFolder(folderPath As String) called on a correct path.
'               - expected : folder cleant
'---------------------------------------------------------------------------------------
'
Public Sub Test_vtkCleanFolder__correct_path()
    Dim retval As String
    Dim fso As New FileSystemObject
    
    vtkCleanFolder (vtkTestPath & "\" & "testFolder0")
    
    mAssert.Should vtkIsFolderEmpty(vtkTestPath & "\" & "testFolder0") = True, "The folder must be empty."
End Sub


'---------------------------------------------------------------------------------------
' Procedures: Test_vtkIsFolderEmpty__*
' Author    : Lucas Vitorino
' Purpose   : Test thte vtkIsFolderEmpty function in different cases. Code is transparent.
'---------------------------------------------------------------------------------------
'
Public Sub Test_vtkIsFolderEmpty__empty_folder()
    mAssert.Should vtkIsFolderEmpty(vtkTestPath & "\" & "testFolder0" & "\" & "testFolder11") = True, "The folder is empty"
End Sub

Public Sub Test_vtkIsFolderEmpty__folder_with_a_subfolder()
    mAssert.Should vtkIsFolderEmpty(vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1" & "\" & "testFolder2") = False, "The folder is not empty : it contains a subfolder"
End Sub

Public Sub Test_vtkIsFolderEmpty__folder_with_a_file()
    mAssert.Should vtkIsFolderEmpty(vtkTestPath & "\" & "testFolder0" & "\" & _
    "testFolder1" & "\" & "testFolder2" & "\" & "testFolder3") = False, _
    "The folder is not empty : it contains a file"
End Sub

Public Sub Test_vtkIsFolderEmpty__folder_with_a_file_and_a_subfolder()
    mAssert.Should vtkIsFolderEmpty(vtkTestPath & "\" & "testFolder0" & "\" & "testFolder1") = False, _
    "The folder is not empty : it contains a file and a subfolder"
End Sub


'---------------------------------------------------------------------------------------
' Procedures: Test_vtkDoesFolderExist__**
' Author    : Lucas Vitorino
' Purpose   : Test the vtkDoesFolderExist function in both the cases : yes and no.
'---------------------------------------------------------------------------------------
'
Public Sub Test_vtkDoesFolderExist__folder_existing()
    mAssert.Should vtkDoesFolderExist(vtkTestPath & "\" & "testFolder0") = True, "Folder exists."
End Sub

Public Sub Test_vtkDoesFolderExist__folder_not_existing()
    mAssert.Should vtkDoesFolderExist(vtkTestPath & "\" & "dummyfolder") = False, "Folder does not exist."
End Sub


