VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkGitFunctionsTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
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
    
    Dim fso As New FileSystemObject
    fso.CreateTextFile ("file.txt")

End Sub

Private Sub ITestCase_TearDown()
    vtkCleanFolder vtkTestPath
End Sub

Public Sub Test_vtkInitializeGit_correctPathExistenceOfDotGitFolder()
   vtkInitializeGit vtkTestPath
   mAssert.Should vtkDoesFolderExist(vtkTestPath & "\.git") = True, ".git folder must exist"
End Sub

Public Sub Test_vtkInitializeGit_correctPathRetVal()
   Dim retVal As Integer
   retVal = vtkInitializeGit(vtkTestPath)
   mAssert.Should retVal = VTK_OK, "Return value must be VTK_OK, it's " & retVal
End Sub

Public Sub Test_vtkInitializeGit_correctPathReInitializeRetVal()
    Dim retVal As Integer
    vtkInitializeGit (vtkTestPath)
    retVal = vtkInitializeGit(vtkTestPath)
    mAssert.Should retVal = VTK_GIT_ALREADY_INITIALIZED_IN_FOLDER, "Return value must be VTK_GIT_ALREADY_INITIALIZED"
End Sub

Public Sub Test_vtkInitializeGit_wrongPathUnexistingFolder()
    Dim retVal As Integer
    retVal = vtkInitializeGit(vtkTestPath & "\no_folder")
    mAssert.Should retVal = VTK_WRONG_FOLDER_PATH, "Return value must be VTK_WRONG_FOLDER_PATH"
End Sub

Public Sub Test_vtkInitializeGit_wrongPathFile()
    Dim retVal As Integer
    retVal = vtkInitializeGit(vtkTestPath & "file.txt")
    mAssert.Should retVal = VTK_WRONG_FOLDER_PATH, "Return value must be VTK_WRONG_FOLDER_PATH"
End Sub

Public Sub Test_vtkInitializeGit_forbiddenPath1()
    Dim retVal As Integer
    retVal = vtkInitializeGit("D:/dummy")
    mAssert.Should retVal = VTK_FORBIDDEN_PARAMETER, "Return value must be VTK_FORBIDDEN_PARAMETER"
End Sub

Public Sub Test_vtkInitializeGit_forbiddenPath2()
    Dim retVal As Integer
    retVal = vtkInitializeGit("DUMMY:/dummy")
    mAssert.Should retVal = VTK_FORBIDDEN_PARAMETER, "Return value must be VTK_FORBIDDEN_PARAMETER, is " & retVal
End Sub

Public Sub Test_vtkInitializeGit_correctPathCustomLogFile()
    'mAssert.Should retVal = VTK_WRONG_FOLDER_PATH, "Return value must be VTK_WRONG_FOLDER_PATH"
End Sub

Public Sub Test_vtkGitConvertWinPath_absolutePathOnCDrive()
    Dim path As String
    Dim convertedPath As String
    path = "C:\test1\test2"
    convertedPath = vtkGitConvertWinPath(path)
    mAssert.Should convertedPath = Chr(34) & "/test1/test2" & Chr(34), "Path must be correctly transformed."
End Sub

Public Sub Test_vtkGitConvertWinPath_absolutePathOnOneLetterDriveNotC()
    Dim path As String
    Dim convertedPath As String
    path = "D:\test1\test2"
    convertedPath = vtkGitConvertWinPath(path)
    mAssert.Should convertedPath = VTK_FORBIDDEN_PARAMETER, "Path is not valid and shouldn't be allowed."
End Sub

Public Sub Test_vtkGitConvertWinPath_absolutePathOnSeveralLettersDrive()
    Dim path As String
    Dim convertedPath As String
    path = "TOTO:\test1\test2"
    convertedPath = vtkGitConvertWinPath(path)
    mAssert.Should convertedPath = VTK_FORBIDDEN_PARAMETER, "Path is not valid and shouldn't be allowed."
End Sub

Public Sub Test_vtkGitConvertWinPath_relativePath()
    Dim path As String
    Dim convertedPath As String
    path = "\test1\test2"
    convertedPath = vtkGitConvertWinPath(path)
    mAssert.Should convertedPath = VTK_FORBIDDEN_PARAMETER, "Path is not valid and shouldn't be allowed."
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_correctPathExistenceOfDotGitFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_correctPathRetVal"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_correctPathReInitializeRetVal"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_wrongPathUnexistingFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_wrongPathFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_forbiddenPath1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_forbiddenPath2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_correctPathCustomLogFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGitConvertWinPath_absolutePathOnCDrive"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGitConvertWinPath_absolutePathOnOneLetterDriveNotC"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGitConvertWinPath_absolutePathOnSeveralLettersDrive"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkGitConvertWinPath_relativePath"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkInitializeGit_correctPathExistenceOfDotGitFolder": Test_vtkInitializeGit_correctPathExistenceOfDotGitFolder
        Case "Test_vtkInitializeGit_correctPathRetVal": Test_vtkInitializeGit_correctPathRetVal
        Case "Test_vtkInitializeGit_correctPathReInitializeRetVal": Test_vtkInitializeGit_correctPathReInitializeRetVal
        Case "Test_vtkInitializeGit_wrongPathUnexistingFolder": Test_vtkInitializeGit_wrongPathUnexistingFolder
        Case "Test_vtkInitializeGit_wrongPathFile": Test_vtkInitializeGit_wrongPathFile
        Case "Test_vtkInitializeGit_forbiddenPath1": Test_vtkInitializeGit_forbiddenPath1
        Case "Test_vtkInitializeGit_forbiddenPath2": Test_vtkInitializeGit_forbiddenPath2
        Case "Test_vtkInitializeGit_correctPathCustomLogFile": Test_vtkInitializeGit_correctPathCustomLogFile
        Case "Test_vtkGitConvertWinPath_absolutePathOnCDrive": Test_vtkGitConvertWinPath_absolutePathOnCDrive
        Case "Test_vtkGitConvertWinPath_absolutePathOnOneLetterDriveNotC": Test_vtkGitConvertWinPath_absolutePathOnOneLetterDriveNotC
        Case "Test_vtkGitConvertWinPath_absolutePathOnSeveralLettersDrive": Test_vtkGitConvertWinPath_absolutePathOnSeveralLettersDrive
        Case "Test_vtkGitConvertWinPath_relativePath": Test_vtkGitConvertWinPath_relativePath
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


