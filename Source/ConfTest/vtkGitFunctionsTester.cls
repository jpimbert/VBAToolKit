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


'---------------------------------------------------------------------------------------
' Procedure : TestvtkInitializeGitFunction
' Author    : Abdelfattah Lahbib
' Purpose   : Test the initialization of a repository in the project root folder
'               - $projectRootFolder\GitLog\logGitInitialize.log must start by "Chr(10) & Initialized"
'---------------------------------------------------------------------------------------
'
Public Sub TestvtkInitializeRepository()
    Dim RetValInitializeGit As String
    Dim RetValFileReader As String
    RetValInitializeGit = vtkInitializeGit()
    RetValFileReader = vtkTextFileReader(vtkTestPath & "\" & vtkTestProjectName & "\GitLog" & "\logGitInitialize.log")
    mAssert.Equals Left(RetValFileReader, 12), Chr(10) & "Initialized", "repository must be initilized"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TestvtkReInitializeLocalRepository
' Author    : Abdelfattah Lahbib
' Purpose   : Test the re-initialization of a repository in a project root folder
'               - $projectRootFolder\GitLog\logGitInitialize.log must start by "Chr(10) & Reinitialized"
'---------------------------------------------------------------------------------------
'
Public Sub TestReInitializeLocalRepository()
    Dim RetValInitializeGit As String
    Dim RetValInitializeGit2 As String
    Dim RetValFileReader As String
    
    RetValInitializeGit = vtkInitializeGit()
    RetValInitializeGit2 = vtkInitializeGit()
 
    RetValFileReader = vtkTextFileReader(vtkTestPath & "\" & vtkTestProjectName & "\GitLog" & "\logGitInitialize.log")
    mAssert.Equals Left(RetValFileReader, 14), Chr(10) & "Reinitialized", "repository must be reinitilized"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestInitializeGitFileIsCreated
' Author    : Abdelfattah Lahbib
' Purpose   : Test the initialization of a repository in the project root folder
'               - $projectRootFolder\GitLog\logGitInitialize.log must exist
'---------------------------------------------------------------------------------------
'
Public Sub TestInitializeGitFileIsCreated()
    Dim RetValInitializeGit As String
    Dim fso As New FileSystemObject
   
    RetValInitializeGit = vtkInitializeGit()
    mAssert.Should fso.FileExists(vtkTestPath & "\" & vtkTestProjectName & "\GitLog\logGitInitialize.log") = True, "log file must be created"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TestInitializeGitFileIsNotEmpty
' Author    : Abdelfattah Lahbib
' Purpose   : Test the initialization of a repository in the project root folder
'               - $projectRootFolder\GitLog\logGitInitialize.log must not be empty
'---------------------------------------------------------------------------------------
'
Public Sub TestInitializeGitFileIsNotEmpty()
    Dim RetValInitializeGit As String
    Dim fso As New FileSystemObject
    
    RetValInitializeGit = vtkInitializeGit()
    
    mAssert.Should IsEmpty(vtkTestPath & "\" & vtkTestProjectName & "\GitLog\logGitInitialize.log") = False, "log file must not be empty"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TestStatusGitFileIsCreated
' Author    : Abdelfattah Lahbib
' Purpose   : Test the vtkStatusGit function
'               - After calling vtkStatusGit, $projectRootFolder\GitLog\logStatus.log must exist
'---------------------------------------------------------------------------------------
'
Public Sub TestStatusGitIsCreated()
   Dim RetValStatusGit As String
   Dim fso As New FileSystemObject
   
   RetValStatusGit = vtkStatusGit()
   
   mAssert.Should fso.FileExists(vtkTestPath & "\" & vtkTestProjectName & "\GitLog\logStatus.log") = True, "status file must be created"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TestStatusGitFileIsCreated
' Author    : Abdelfattah Lahbib
' Purpose   : Test the vtkStatusGit function
'               - After calling vtkStatusGit, $projectRootFolder\GitLog\logStatus.log must not be empty
'---------------------------------------------------------------------------------------
'
Public Sub TestStatusGitFileNotEmpty()
    Dim RetValStatusGit As String
    Dim fso As New FileSystemObject
    
    RetValStatusGit = vtkStatusGit()
 
    mAssert.Should IsEmpty(vtkTestPath & "\" & vtkTestProjectName & "\GitLog\logStatus.log") = False, "log file must not be empty"
End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestvtkInitializeRepository"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReInitializeLocalRepository"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestInitializeGitFileIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestInitializeGitFileIsNotEmpty"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestStatusGitIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestStatusGitFileNotEmpty"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestvtkInitializeRepository": TestvtkInitializeRepository
        Case "TestReInitializeLocalRepository": TestReInitializeLocalRepository
        Case "TestInitializeGitFileIsCreated": TestInitializeGitFileIsCreated
        Case "TestInitializeGitFileIsNotEmpty": TestInitializeGitFileIsNotEmpty
        Case "TestStatusGitIsCreated": TestStatusGitIsCreated
        Case "TestStatusGitFileNotEmpty": TestStatusGitFileNotEmpty
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


