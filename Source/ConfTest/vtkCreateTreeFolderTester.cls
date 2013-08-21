VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkCreateTreeFolderTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
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
End Sub

Private Sub ITestCase_TearDown()
'   Delete Folder Tree
    vtkCleanFolder vtkTestPath
End Sub

Public Sub TestCreateTreeFolderDoesNothingWhenPathDoesntExist()
    ' The Create Tree Folder function must does nothing and return error when path doesn't exist
    Dim InexistentPath As String
    Dim error As Integer: error = 0
    
    On Error GoTo errorHandler
    
    InexistentPath = vtkTestPath & "\InexistentFolder"
    vtkCreateTreeFolder InexistentPath & "\" & vtkTestProjectName
    mAssert.Should Dir(InexistentPath & "\" & vtkTestProjectName, vbDirectory) = "", "The project main folder must not exist"
    mAssert.Equals error, 76, "The project creation must fail"
    
    On Error GoTo 0
    Exit Sub
    
errorHandler:
    error = Err.Number
    Resume Next
End Sub

Public Sub TestCreateTreeFolderAlreadyExisting()
    Dim error As Integer: error = 0
    
    On Error GoTo errorHandler
    
    MkDir vtkTestPath & "\treeroot"
    vtkCreateTreeFolder vtkTestPath & "\treeroot"
    mAssert.Equals error, 75, "The project creation must fail"
    
    On Error GoTo 0
    Exit Sub
    
errorHandler:
    error = Err.Number
    Resume Next
    
End Sub

Public Sub TestMainFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named like the project in the path
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName, vbDirectory) <> "", "The project main folder must exist"
    mAssert.Equals returnvalue, VTK_OK, "The project creation must succeed"
End Sub

Public Sub TestProjectFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named Project
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Project", vbDirectory) <> "", "The Project folder must exist"
    mAssert.Equals returnvalue, VTK_OK, "The project creation must succeed"
End Sub

Public Sub TestSourceFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named Source
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Source", vbDirectory) <> "", "The Source folder must exist"
    mAssert.Equals returnvalue, VTK_OK, "The project creation must succeed"
End Sub

Public Sub TestTestsFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named Tests
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Tests", vbDirectory) <> "", "The Tests folder must exist"
    mAssert.Equals returnvalue, VTK_OK, "The project creation must succeed"
End Sub

Public Sub TestConfProdFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named ConfProd
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\ConfProd", vbDirectory) <> "", "source\ConfProd folder must exist"
    mAssert.Equals returnvalue, VTK_OK, "The project creation must succeed"
End Sub

Public Sub TestConfTestFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named ConfProd
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\ConfTest", vbDirectory) <> "", "The source\ConfTest folder must exist"
    mAssert.Equals returnvalue, VTK_OK, "The project creation must succeed"
End Sub

Public Sub TestVbaUnitFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named ConfProd
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\VbaUnit", vbDirectory) <> "", "The source\VbaUnit folder must exist"
    mAssert.Equals returnvalue, VTK_OK, "The project creation must succeed"
End Sub

Public Sub TestGitLogIsCreated()
    ' The Create Tree Folder function must create a new folder named GitLog
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\GitLog", vbDirectory) <> "", "The ..\GitLog folder must exist"
    mAssert.Equals returnvalue, VTK_OK, "The project creation must succeed"
End Sub

Public Sub TestDeliveryFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named Delivery
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Delivery", vbDirectory) <> "", "The \Delivery folder must exist"
    mAssert.Equals returnvalue, VTK_OK, "The project creation must succeed"
End Sub

Public Sub Test_vtkCreateTreeFolderWithGitIgnore_correctPathReturnValue()
    Dim retval As Integer
    retval = vtkCreateTreeFolderWithGitIgnore(vtkTestPath & "\treeroot")
    mAssert.Equals retval, VTK_OK
End Sub

Public Sub Test_vtkCreateTreeFolderWithGitIgnore_DoesNothingWhenPathDoesntExist()
    ' The Create Tree Folder function must do nothing and return error when path doesn't exist
    Dim InexistentPath As String
    Dim returnvalue As Long
    Dim error As Integer: error = 0
    InexistentPath = vtkTestPath & "\InexistentFolder"
    
    On Error GoTo errorHandler
    
    returnvalue = vtkCreateTreeFolderWithGitIgnore(InexistentPath & "\treeroot")
    mAssert.Should Dir(InexistentPath & "\treeroot", vbDirectory) = "", "The project main folder must not exist"
    mAssert.Equals error, 76, "The creation of the folder must fail"
    
    On Error GoTo 0
    Exit Sub
    
errorHandler:
    error = Err.Number
    Resume Next
End Sub

Public Sub Test_vtkCreateTreeFolderWithGitIgnore_AlreadyExisting()
    ' The Create Tree Folder function must do nothing when the root folder already exists
    Dim InexistentPath As String
    Dim error As Integer: error = 0
    
    On Error GoTo errorHandler
    
    MkDir vtkTestPath & "\treeroot"
    vtkCreateTreeFolderWithGitIgnore vtkTestPath & "\treeroot"
    mAssert.Should Dir(vtkTestPath & "\treeroot", vbDirectory) <> "", "The existing project main folder must exist"
    mAssert.Equals error, 75, "The project creation must fail"
    
    On Error GoTo 0
    Exit Sub
    
errorHandler:
    error = Err.Number
    Resume Next
End Sub

Public Sub Test_vtkCreateTreeFolderWithGitIgnore_correctPathFileExistsInRoot()
    Dim fso As New FileSystemObject
    vtkCreateTreeFolderWithGitIgnore (vtkTestPath & "\treeroot")
    mAssert.Equals fso.FileExists(vtkTestPath & "\treeroot\.gitignore"), True
End Sub

Public Sub Test_vtkCreateTreeFolderWithGitIgnore_correctPathFileContentInRoot()
    Dim contentString As String
    vtkCreateTreeFolderWithGitIgnore (vtkTestPath & "\treeroot")
    contentString = vtkTextFileReader(vtkTestPath & "\treeroot\.gitignore")
    mAssert.Equals Left(contentString, 9), Chr(10) & "# Ignore"
End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateTreeFolderDoesNothingWhenPathDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateTreeFolderAlreadyExisting"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestMainFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSourceFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestTestsFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfProdFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfTestFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVbaUnitFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGitLogIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDeliveryFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateTreeFolderWithGitIgnore_correctPathReturnValue"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateTreeFolderWithGitIgnore_DoesNothingWhenPathDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateTreeFolderWithGitIgnore_AlreadyExisting"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateTreeFolderWithGitIgnore_correctPathFileExistsInRoot"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateTreeFolderWithGitIgnore_correctPathFileContentInRoot"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestCreateTreeFolderDoesNothingWhenPathDoesntExist": TestCreateTreeFolderDoesNothingWhenPathDoesntExist
        Case "TestCreateTreeFolderAlreadyExisting": TestCreateTreeFolderAlreadyExisting
        Case "TestMainFolderIsCreated": TestMainFolderIsCreated
        Case "TestProjectFolderIsCreated": TestProjectFolderIsCreated
        Case "TestSourceFolderIsCreated": TestSourceFolderIsCreated
        Case "TestTestsFolderIsCreated": TestTestsFolderIsCreated
        Case "TestConfProdFolderIsCreated": TestConfProdFolderIsCreated
        Case "TestConfTestFolderIsCreated": TestConfTestFolderIsCreated
        Case "TestVbaUnitFolderIsCreated": TestVbaUnitFolderIsCreated
        Case "TestGitLogIsCreated": TestGitLogIsCreated
        Case "TestDeliveryFolderIsCreated": TestDeliveryFolderIsCreated
        Case "Test_vtkCreateTreeFolderWithGitIgnore_correctPathReturnValue": Test_vtkCreateTreeFolderWithGitIgnore_correctPathReturnValue
        Case "Test_vtkCreateTreeFolderWithGitIgnore_DoesNothingWhenPathDoesntExist": Test_vtkCreateTreeFolderWithGitIgnore_DoesNothingWhenPathDoesntExist
        Case "Test_vtkCreateTreeFolderWithGitIgnore_AlreadyExisting": Test_vtkCreateTreeFolderWithGitIgnore_AlreadyExisting
        Case "Test_vtkCreateTreeFolderWithGitIgnore_correctPathFileExistsInRoot": Test_vtkCreateTreeFolderWithGitIgnore_correctPathFileExistsInRoot
        Case "Test_vtkCreateTreeFolderWithGitIgnore_correctPathFileContentInRoot": Test_vtkCreateTreeFolderWithGitIgnore_correctPathFileContentInRoot
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


