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
     MkDir vtkTestPath & "\" & vtkTestProjectName
     MkDir vtkTestPath & "\" & vtkTestProjectName & "\Project"
     MkDir vtkTestPath & "\" & vtkTestProjectName & "\GitLog"

     Workbooks.Add.SaveAs (vtkTestPath & "\" & vtkTestProjectName & "\Project\" & vtkTestProjectName), FileFormat:=(52) '52 is xlsm format

End Sub

Private Sub ITestCase_TearDown()
    Dim fso As New FileSystemObject
    Application.DisplayAlerts = False
    On Error Resume Next
    'Close Created WorkBook
    Workbooks(vtkTestProjectName & ".xlsm").Close
   'Delete Created WorkBook
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Project\" & vtkTestProjectName & ".xlsm"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Project"
    fso.DeleteFolder (vtkTestPath & "\" & vtkTestProjectName & "\.git")
    Kill vtkTestPath & "\" & vtkTestProjectName & "\GitLog\*"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\GitLog"
    Kill vtkTestPath & "\" & vtkTestProjectName & "\*"
    RmDir vtkTestPath & "\" & vtkTestProjectName
    Application.DisplayAlerts = True
End Sub

Public Sub TestvtkVerifyEnvirGitVarFunction()
    'verify that envirenement  var exist
    Dim RetValVerifEnvVar As String
    RetValVerifEnvVar = vtkVerifyEnvirGitVar()
    mAssert.Equals RetValVerifEnvVar, "", "envirenement variable must be setted"

End Sub


Public Sub TestvtkInitializeGitFunction()
'test that the repository is initilized for the first time
    Dim RetValInitializeGit As String
    Dim RetValFileReader As String
   RetValInitializeGit = vtkInitializeGit()
   RetValFileReader = vtkTextFileReader(vtkTestPath & "\" & vtkTestProjectName & "\GitLog" & "\logGitInitialize.log")
   mAssert.Equals Left(RetValFileReader, 12), Chr(10) & "Initialized", "repository must be initilized"

End Sub
Public Sub TestReInitializeLocalRepository()
'test that repository is initilized tow time
    Dim RetValInitializeGit As String
    Dim RetValInitializeGit2 As String
    Dim RetValFileReader As String
    
   RetValInitializeGit = vtkInitializeGit()
   RetValInitializeGit2 = vtkInitializeGit()
   
   RetValFileReader = vtkTextFileReader(vtkTestPath & "\" & vtkTestProjectName & "\GitLog" & "\logGitInitialize.log")
    mAssert.Equals Left(RetValFileReader, 14), Chr(10) & "Reinitialized", "repository must be reinitilized"
End Sub
Public Sub TestInitializeGitFileIsCreated()
'test that the log file exist
 Dim RetValInitializeGit As String
 Dim fso As New FileSystemObject
   
   RetValInitializeGit = vtkInitializeGit()
   mAssert.Should fso.FileExists(vtkTestPath & "\" & vtkTestProjectName & "\GitLog\logGitInitialize.log") = True, "log file must be created"
End Sub
Public Sub TestInitializeGitFileNotEmpty()
' test that the log file is not empty
    Dim RetValInitializeGit As String
    Dim fso As New FileSystemObject
   RetValInitializeGit = vtkInitializeGit()
 
 mAssert.Should IsEmpty(vtkTestPath & "\" & vtkTestProjectName & "\GitLog\logGitInitialize.log") = False, "log file must not be empty"
End Sub


Public Sub TestStatusGitIsCreated()
 'test that status file is created
   Dim RetValStatusGit As String
   Dim fso As New FileSystemObject
   
   RetValStatusGit = vtkStatusGit()
 
   mAssert.Should fso.FileExists(vtkTestPath & "\" & vtkTestProjectName & "\GitLog\logStatus.log") = True, "status file must be created"
End Sub
Public Sub TestStatusGitFileNotEmpty()
'test that status file is not empty
    Dim RetValStatusGit As String
    Dim fso As New FileSystemObject
    RetValStatusGit = vtkStatusGit()
 
 mAssert.Should IsEmpty(vtkTestPath & "\" & vtkTestProjectName & "\GitLog\logStatus.log") = False, "log file must not be empty"
End Sub
Public Sub TestvtkCreateFileInDirectoryFunction()
'test the function that create log files
    Dim retVal As String
    Dim fso As New FileSystemObject
   
   retVal = vtkCreateFileInDirectory(vtkTestProjectName & ".txt", vtkTestPath & "\" & vtkTestProjectName & "\GitLog\")
 
   mAssert.Should fso.FileExists(vtkTestPath & "\" & vtkTestProjectName & "\GitLog\" & vtkTestProjectName & ".txt") = True, "log file must be created"
End Sub
Public Sub TestVtkFileReaderFunction()
 'test functions that read files
    Dim retVal As String
    Dim fso As New FileSystemObject
   
   retVal = vtkCreateFileInDirectory(vtkTestProjectName & ".txt", vtkTestPath & "\" & vtkTestProjectName & "\GitLog\")
    
   mAssert.Should IsEmpty(vtkTestPath & "\" & vtkTestProjectName & "\GitLog\" & vtkTestProjectName & ".txt") = False, "reader must work"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestvtkVerifyEnvirGitVarFunction"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestvtkInitializeGitFunction"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReInitializeLocalRepository"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestInitializeGitFileIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestInitializeGitFileNotEmpty"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestStatusGitIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestStatusGitFileNotEmpty"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestvtkCreateFileInDirectoryFunction"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVtkFileReaderFunction"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestvtkVerifyEnvirGitVarFunction": TestvtkVerifyEnvirGitVarFunction
        Case "TestvtkInitializeGitFunction": TestvtkInitializeGitFunction
        Case "TestReInitializeLocalRepository": TestReInitializeLocalRepository
        Case "TestInitializeGitFileIsCreated": TestInitializeGitFileIsCreated
        Case "TestInitializeGitFileNotEmpty": TestInitializeGitFileNotEmpty
        Case "TestStatusGitIsCreated": TestStatusGitIsCreated
        Case "TestStatusGitFileNotEmpty": TestStatusGitFileNotEmpty
        Case "TestvtkCreateFileInDirectoryFunction": TestvtkCreateFileInDirectoryFunction
        Case "TestVtkFileReaderFunction": TestVtkFileReaderFunction
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


