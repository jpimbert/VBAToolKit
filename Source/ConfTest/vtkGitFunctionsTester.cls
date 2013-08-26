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
    On Error GoTo Test_vtkInitializeGit_correctPathExistenceOfDotGitFolder_Error

   vtkInitializeGit vtkTestPath
   mAssert.Should vtkDoesFolderExist(vtkTestPath & "\.git") = True, ".git folder must exist"

    On Error GoTo 0
    Exit Sub

Test_vtkInitializeGit_correctPathExistenceOfDotGitFolder_Error:
    Debug.Print "Error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_correctPathExistenceOfDotGitFolder of Class Module vtkGitFunctionsTester"
    Resume Next
End Sub

Public Sub Test_vtkInitializeGit_correctPathRetVal()
    Dim retval As Integer
    On Error GoTo Test_vtkInitializeGit_correctPathRetVal_Error

    retval = vtkInitializeGit(vtkTestPath)
    mAssert.Equals retval, VTK_OK

    On Error GoTo 0
    Exit Sub

Test_vtkInitializeGit_correctPathRetVal_Error:
    Debug.Print "Error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_correctPathRetVal of Class Module vtkGitFunctionsTester"
    Resume Next
End Sub

Public Sub Test_vtkInitializeGit_correctPathReInitializeRetVal()
    Dim error As Integer: error = 0
    
    On Error GoTo errHandler
    
    vtkInitializeGit (vtkTestPath)
    vtkInitializeGit (vtkTestPath)
    
    mAssert.Equals error, VTK_GIT_ALREADY_INITIALIZED_IN_FOLDER
    
    On Error GoTo 0
    Exit Sub
    
errHandler:
    'Debug.Print "EXPECTED Error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_correctPathReInitializeRetVal of Class Module vtkGitFunctionsTester"
    error = Err.number
    Resume Next
End Sub

Public Sub Test_vtkInitializeGit_wrongPathUnexistingFolder()
    Dim error As Integer: error = 0
    
    On Error GoTo errHandler
    
    vtkInitializeGit (vtkTestPath & "\no_folder")
    
    mAssert.Equals error, VTK_WRONG_FOLDER_PATH
    
    On Error GoTo 0
    Exit Sub
    
errHandler:
    'Debug.Print "EXPECTED Error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_wrongPathUnexistingFolder of Class Module vtkGitFunctionsTester"
    error = Err.number
    Resume Next
    
End Sub

Public Sub Test_vtkInitializeGit_wrongPathFile()
    Dim error As Integer: error = 0
    
    On Error GoTo errHandler
    
    vtkInitializeGit (vtkTestPath & "file.txt")
    
    mAssert.Equals error, VTK_WRONG_FOLDER_PATH
    
    On Error GoTo 0
    Exit Sub
    
errHandler:
    'Debug.Print "EXPECTED Error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_wrongPathFile of Class Module vtkGitFunctionsTester"
    error = Err.number
    Resume Next
End Sub

Public Sub Test_vtkInitializeGit_forbiddenPath1()
    Dim error As Integer: error = 0
    
    On Error GoTo errHandler
    
    vtkInitializeGit ("D:/dummy")
    
    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER
    
    On Error GoTo 0
    Exit Sub
    
errHandler:
    'Debug.Print "EXPECTED Error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_forbiddenPath1 of Class Module vtkGitFunctionsTester"
    error = Err.number
    Resume Next
    
End Sub

Public Sub Test_vtkInitializeGit_forbiddenPath2()
    Dim error As Integer: error = 0
    
    On Error GoTo errHandler
    
    vtkInitializeGit ("DUMMY:/dummy")
    
    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER
    
    On Error GoTo 0
    Exit Sub
    
errHandler:
    'Debug.Print "EXPECTED Error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_forbiddenPath2 of Class Module vtkGitFunctionsTester"
    error = Err.number
    Resume Next
    
End Sub

Public Sub Test_vtkInitializeGit_correctPathLogFileIsDeleted()
    Dim fso As New Scripting.FileSystemObject
    On Error GoTo Test_vtkInitializeGit_correctPathLogFileIsDeleted_Error

    vtkInitializeGit vtkTestPath
    mAssert.Equals fso.FileExists(vtkTestPath & "\initialize.log"), False

    On Error GoTo 0
    Exit Sub

Test_vtkInitializeGit_correctPathLogFileIsDeleted_Error:
    Debug.Print "Error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_correctPathLogFileIsDeleted of Class Module vtkGitFunctionsTester"
    Resume Next
End Sub

Public Sub Test_vtkInitializeGit_correctPathCustomLogFileReturnValue()
    Dim retval As Integer
    On Error GoTo Test_vtkInitializeGit_correctPathCustomLogFileReturnValue_Error

    retval = vtkInitializeGit(vtkTestPath, "initialize.log")
    mAssert.Equals retval, VTK_OK

    On Error GoTo 0
    Exit Sub

Test_vtkInitializeGit_correctPathCustomLogFileReturnValue_Error:
    Debug.Print "Error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_correctPathCustomLogFileReturnValue of Class Module vtkGitFunctionsTester"
    Resume Next
End Sub

Public Sub Test_vtkInitializeGit_correctPathCustomLogFileExists()
    Dim fso As New FileSystemObject
    On Error GoTo Test_vtkInitializeGit_correctPathCustomLogFileExists_Error

    vtkInitializeGit vtkTestPath, "initialize.log"
    mAssert.Equals fso.FileExists(vtkTestPath & "\initialize.log"), True

    On Error GoTo 0
    Exit Sub

Test_vtkInitializeGit_correctPathCustomLogFileExists_Error:
    Debug.Print "Error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_correctPathCustomLogFileExists of Class Module vtkGitFunctionsTester"
    Resume Next
End Sub

Public Sub Test_vtkGitConvertWinPath_absolutePathOnCDrive()
    Dim path As String
    Dim convertedPath As String
    On Error GoTo Test_vtkGitConvertWinPath_absolutePathOnCDrive_Error

    path = "C:\test1\test2"
    convertedPath = vtkGitConvertWinPath(path)
    mAssert.Equals convertedPath, Chr(34) & "/test1/test2" & Chr(34), "Path must be correctly transformed."

    On Error GoTo 0
    Exit Sub

Test_vtkGitConvertWinPath_absolutePathOnCDrive_Error:
    Debug.Print "Error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkGitConvertWinPath_absolutePathOnCDrive of Class Module vtkGitFunctionsTester"
    Resume Next
End Sub

Public Sub Test_vtkGitConvertWinPath_absolutePathOnOneLetterDriveNotC()
    Dim error As Integer: error = 0
    Dim path As String
    Dim convertedPath As String
    
    On Error GoTo errHandler
    
    path = "D:\test1\test2"
    convertedPath = vtkGitConvertWinPath(path)
    
    mAssert.Equals error, VTK_FORBIDDEN_PARAMETER
    
    On Error GoTo 0
    Exit Sub
    
errHandler:
    'Debug.Print "EXPECTED Error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_wrongPathFile of Class Module vtkGitFunctionsTester"
    error = Err.number
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
    
    On Error GoTo 0
    Exit Sub
    
errHandler:
    'Debug.Print "EXPECTED Error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGitConvertWinPath_absolutePathOnSeveralLettersDrive of Class Module vtkGitFunctionsTester"
    error = Err.number
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
    
    On Error GoTo 0
    Exit Sub
    
errHandler:
    'Debug.Print "EXPECTED Error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkGitConvertWinPath_relativePath of Class Module vtkGitFunctionsTester"
    error = Err.number
    Resume Next
    
End Sub

Public Sub Test_vtkInitializeGit_correctPathGitExcludeContentIsRight()

    Dim fileContent As String
    Dim expectedContent As String
    Dim fso As New FileSystemObject

    On Error GoTo Test_vtkInitializeGit_correctPathGitExcludeContentIsRight_Error

    vtkInitializeGit vtkTestPath
    
    fileContent = vtkTextFileReader(vtkTestPath & "\.git\info\exclude")
    
    expectedContent = _
    "# Ignore the content of the Tests and GitLog folders" & Chr(10) & _
    "/Tests/*" & Chr(10) & _
    "/GitLog/*" & Chr(10) & _
    Chr(10) & _
    "# Ignore the temporary Excel files" & Chr(10) & _
    "/Project/~*" & Chr(10) & _
    Chr(10) & _
    "# Ignore the delivery Excel files" & Chr(10) & _
    "/Delivery/*.xlsm"
    
    mAssert.Equals StrComp(fileContent, expectedContent), True

    On Error GoTo 0
    Exit Sub

Test_vtkInitializeGit_correctPathGitExcludeContentIsRight_Error:
    Debug.Print "Error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkInitializeGit_correctPathGitExcludeContentIsRight of Class Module vtkGitFunctionsTester"
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
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkInitializeGit_correctPathGitExcludeContentIsRight"
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
        Case "Test_vtkInitializeGit_correctPathGitExcludeContentIsRight": Test_vtkInitializeGit_correctPathGitExcludeContentIsRight
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


