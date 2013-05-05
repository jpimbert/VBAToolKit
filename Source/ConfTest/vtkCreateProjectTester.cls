VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkCreateProjectTester"
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
    
    'allow to iterate over all files in a directory
    Dir (vtkTestPath)
    'to not display messagebox how ask to save project
    Application.DisplayAlerts = False
    On Error Resume Next
    'Close Created WorkBook
    Workbooks(vtkTestProjectName & ".xlsm").Close
    'Delete Created WorkBook
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Project\" & vtkTestProjectName & ".xlsm"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\ConfProd"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\ConfTest"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Tests"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Project"
    RmDir vtkTestPath & "\" & vtkTestProjectName
    Application.DisplayAlerts = True
End Sub


Private Sub ITestCase_TearDown()
    'allow to iterate over all files in a directory
    Dir (vtkTestPath)
    'to not display messagebox how ask to save project
    Application.DisplayAlerts = False
    On Error Resume Next
    'Close Created WorkBook
    Workbooks(vtkTestProjectName & ".xlsm").Close
    'Delete Created WorkBook
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Project\" & vtkTestProjectName & ".xlsm"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\ConfProd"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\ConfTest"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Tests"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Project"
    RmDir vtkTestPath & "\" & vtkTestProjectName
    Application.DisplayAlerts = True

End Sub

Public Sub TestCreateProjectDoesNothingWhenPathDoesntExist()
    ' The createProject must does nothing when path doesn't exist
    Dim InexistentPath As String
    Dim returnValue As Long
    InexistentPath = vtkTestPath & "\InexistentFolder"
    returnValue = vtkCreateProject(path:=InexistentPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(InexistentPath & "\" & vtkTestProjectName, vbDirectory) = "", "The project main folder must exist"
    mAssert.Equals returnValue, 76, "The project creation must fail"
End Sub

Public Sub TestCreateProjectAlreadyExisting()
    ' The createProject must does nothing when project already exists
    Dim returnValue As Long
    MkDir vtkTestPath & "\" & vtkTestProjectName 'must respect order mkdir before call vtkcreateproject
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName, vbDirectory) <> "", "The existing project main folder must exist"
    mAssert.Equals returnValue, 75, "The project creation must fail"
End Sub

Public Sub TestMainFolderIsCreated()
    ' The createProject subroutine must create a new folder named like the project in the path
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName, vbDirectory) <> "", "The project main folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub

Public Sub TestProjectFolderIsCreated()
    ' The createProject subroutine must create a new folder named Project
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Project", vbDirectory) <> "", "The Project folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub

Public Sub TestSourceFolderIsCreated()
    ' The createProject subroutine must create a new folder named Source
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Source", vbDirectory) <> "", "The Source folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub

Public Sub TestTestsFolderIsCreated()
    ' The createProject subroutine must create a new folder named Tests
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Tests", vbDirectory) <> "", "The Tests folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub
Public Sub TestConfProdFolderIsCreated()
    ' The createProject subroutine must create a new folder named ConfProd
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\ConfProd", vbDirectory) <> "", "source\ConfProd folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub
Public Sub TestConfTestFolderIsCreated()
    ' The createProject subroutine must create a new folder named ConfProd
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\ConfTest", vbDirectory) <> "", "The source\ConfTest folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub
Public Sub TestVbaUnitFolderIsCreated()
    ' The createProject subroutine must create a new folder named ConfProd
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\VbaUnit", vbDirectory) <> "", "The source\VbaUnit folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub
Public Sub TestXlsmFileCreated()
'test that xlsm file is created in the right path
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & ".xlsm", vbDirectory) <> "", "the workbook must be created"
End Sub
Public Sub TestXlsmFileIsNamedcorrectly()
'test that xlsm is correctly renamed
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Workbooks(vtkTestProjectName & ".xlsm").VBProject.name = vtkTestProjectName, "the workbook name must be like " & vtkTestProjectName
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateProjectDoesNothingWhenPathDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateProjectAlreadyExisting"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestMainFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSourceFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestTestsFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfProdFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfTestFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVbaUnitFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestXlsmFileCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestXlsmFileIsNamedcorrectly"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestCreateProjectDoesNothingWhenPathDoesntExist": TestCreateProjectDoesNothingWhenPathDoesntExist
        Case "TestCreateProjectAlreadyExisting": TestCreateProjectAlreadyExisting
        Case "TestMainFolderIsCreated": TestMainFolderIsCreated
        Case "TestProjectFolderIsCreated": TestProjectFolderIsCreated
        Case "TestSourceFolderIsCreated": TestSourceFolderIsCreated
        Case "TestTestsFolderIsCreated": TestTestsFolderIsCreated
        Case "TestConfProdFolderIsCreated": TestConfProdFolderIsCreated
        Case "TestConfTestFolderIsCreated": TestConfTestFolderIsCreated
        Case "TestVbaUnitFolderIsCreated": TestVbaUnitFolderIsCreated
        Case "TestXlsmFileCreated": TestXlsmFileCreated
        Case "TestXlsmFileIsNamedcorrectly": TestXlsmFileIsNamedcorrectly
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

