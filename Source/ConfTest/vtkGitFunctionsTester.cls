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
   Dim retval As Integer
   retval = vtkInitializeGit(vtkTestPath)
   mAssert.Equals retval, VTK_OK
End Sub

Public Sub Test_vtkInitializeGit_correctPathReInitializeRetVal()
    Dim error As Integer: error = 0
    
    On Error GoTo errHandler
    
    vtkInitializeGit (vtkTestPath)
    vtkInitializeGit (vtkTestPath)
    
    mAssert.Equals error, VTK_GIT_ALREADY_INITIALIZED_IN_FOLDER
    
errHandler:
    error = Err.Number
    Resume Next
End Sub

Public Sub Test_vtkInitializeGit_wrongPathUnexistingFolder()
    Dim error As Integer: error = 0
    
    On Error GoTo errHandler
    
    vtkInitializeGit (vtkTestPath & "\no_folder")
    
    mAssert.Equals error, VTK_WRONG_FOLDER_PATH
    
errHandler:
    error = Err.Number
    Resume Next
    
End Sub

Public Sub Test_vtkInitializeGit_wrongPathFile()
    Dim error As Integer: error = 0
    
    On Error GoTo errHandler
    
    vtkInitializeGit (vtkTestPath & "file.txt")
    
    mAssert.Equals error, VTK_WRONG_FOLDER_PATH
    
errHandler:
    error = Err.Number
    Resume Next
End Sub

Public Sub Test_vtkInitializeGit_forbiddenPath1()
    Dim error As Integer: error = 0
    
    On Error GoTo errHandler
    
    vtkInitializeGit ("D:/dummy")
    
    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER
    
errHandler:
    error = Err.Number
    Resume Next
    
End Sub

Public Sub Test_vtkInitializeGit_forbiddenPath2()
    Dim error As Integer: error = 0
    
    On Error GoTo errHandler
    
    vtkInitializeGit ("DUMMY:/dummy")
    
    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER
    
errHandler:
    error = Err.Number
    Resume Next
    
End Sub

Public Sub Test_vtkInitializeGit_correctPathLogFileIsDeleted()
    Dim fso As New Scripting.FileSystemObject
    vtkInitializeGit vtkTestPath
    mAssert.Equals fso.FileExists(vtkTestPath & "\initialize.log"), False
End Sub

Public Sub Test_vtkInitializeGit_correctPathCustomLogFileReturnValue()
    Dim retval As Integer
    retval = vtkInitializeGit(vtkTestPath, "initialize.log")
    mAssert.Equals retval, VTK_OK
End Sub

Public Sub Test_vtkInitializeGit_correctPathCustomLogFileExists()
    Dim fso As New FileSystemObject
    vtkInitializeGit vtkTestPath, "initialize.log"
    mAssert.Equals fso.FileExists(vtkTestPath & "\initialize.log"), True
End Sub

Public Sub Test_vtkGitConvertWinPath_absolutePathOnCDrive()
    Dim path As String
    Dim convertedPath As String
    path = "C:\test1\test2"
    convertedPath = vtkGitConvertWinPath(path)
    mAssert.Equals convertedPath, Chr(34) & "/test1/test2" & Chr(34), "Path must be correctly transformed."
End Sub

Public Sub Test_vtkGitConvertWinPath_absolutePathOnOneLetterDriveNotC()
    Dim error As Integer: error = 0
    Dim path As String
    Dim convertedPath As String
    
    On Error GoTo errHandler
    
    path = "D:\test1\test2"
    convertedPath = vtkGitConvertWinPath(path)
    
    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER
    
errHandler:
    error = Err.Number
    Resume Next
    
End Sub

Public Sub Test_vtkGitConvertWinPath_absolutePathOnSeveralLettersDrive()
    Dim error As Integer: error = 0
    Dim path As String
    Dim convertedPath As String
    
    On Error GoTo errHandler
    
    path = "TOTO:\test1\test2"
    convertedPath = vtkGitConvertWinPath(path)
    
    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER
    
errHandler:
    error = Err.Number
    Resume Next
    
End Sub

Public Sub Test_vtkGitConvertWinPath_relativePath()
    Dim error As Integer: error = 0
    Dim path As String
    Dim convertedPath As String
    
    On Error GoTo errHandler
    
    path = "\test1\test2"
    convertedPath = vtkGitConvertWinPath(path)
    
    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER
    
errHandler:
    error = Err.Number
    Resume Next
    
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
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_correctPathLogFileIsDeleted"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_correctPathCustomLogFileReturnValue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_correctPathCustomLogFileExists"
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
        Case "Test_vtkInitializeGit_correctPathLogFileIsDeleted": Test_vtkInitializeGit_correctPathLogFileIsDeleted
        Case "Test_vtkInitializeGit_correctPathCustomLogFileReturnValue": Test_vtkInitializeGit_correctPathCustomLogFileReturnValue
        Case "Test_vtkInitializeGit_correctPathCustomLogFileExists": Test_vtkInitializeGit_correctPathCustomLogFileExists
        Case "Test_vtkGitConvertWinPath_absolutePathOnCDrive": Test_vtkGitConvertWinPath_absolutePathOnCDrive
        Case "Test_vtkGitConvertWinPath_absolutePathOnOneLetterDriveNotC": Test_vtkGitConvertWinPath_absolutePathOnOneLetterDriveNotC
        Case "Test_vtkGitConvertWinPath_absolutePathOnSeveralLettersDrive": Test_vtkGitConvertWinPath_absolutePathOnSeveralLettersDrive
        Case "Test_vtkGitConvertWinPath_relativePath": Test_vtkGitConvertWinPath_relativePath
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


