VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkCreateProjectTester"
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
    
    Dim name  As String 'for file .xla
    name = "TestProject" 'initialize it
    
   ' Delete all files and folder in the Tests Directory
    Dir (vtkTestPath)       ' Make sure to be out of the folder to clean
    On Error Resume Next    ' If there is no files here, this is not an error
    
   
    Workbooks(name & ".xls").Close
    Kill vtkTestPath & "\" & name & "\" & "Project" & "\" & name & ".xls"
    RmDir vtkTestPath & "\TestProject\Source\ConfProd"
    RmDir vtkTestPath & "\TestProject\Source\ConfTest"
    RmDir vtkTestPath & "\TestProject\Source\VbaUnit"
    RmDir vtkTestPath & "\TestProject\Tests"
    RmDir vtkTestPath & "\TestProject\Source"
    RmDir vtkTestPath & "\TestProject\Project"
    RmDir vtkTestPath & "\TestProject"
    
 
    
End Sub

Private Sub ITestCase_TearDown()

End Sub

Public Sub TestCreateProjectDoesNothingWhenPathDoesntExist()
    ' The createProject must does nothing when path doesn't exist
    Dim InexistentPath As String
    Dim returnValue As Long
    InexistentPath = vtkTestPath & "\InexistentFolder"
    returnValue = vtkCreateProject(path:=InexistentPath, name:="TestProject", displayError:=False)
    mAssert.Should Dir(InexistentPath & "\TestProject", vbDirectory) = "", "The project main folder must exist"
    mAssert.Equals returnValue, 76, "The project creation must fail"
End Sub

Public Sub TestCreateProjectAlreadyExisting()
    ' The createProject must does nothing when project already exists
    Dim returnValue As Long
    MkDir vtkTestPath & "\TestProject"
    returnValue = vtkCreateProject(path:=vtkTestPath, name:="TestProject", displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\TestProject", vbDirectory) <> "", "The existing project main folder must exist"
    mAssert.Equals returnValue, 75, "The project creation must fail"
End Sub

Public Sub TestMainFolderIsCreated()
    ' The createProject subroutine must create a new folder named like the project in the path
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:="TestProject", displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\TestProject", vbDirectory) <> "", "The project main folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub

Public Sub TestProjectFolderIsCreated()
    ' The createProject subroutine must create a new folder named Project
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:="TestProject", displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\TestProject\Project", vbDirectory) <> "", "The Project folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub

Public Sub TestSourceFolderIsCreated()
    ' The createProject subroutine must create a new folder named Source
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:="TestProject", displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\TestProject\Source", vbDirectory) <> "", "The Source folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub

Public Sub TestTestsFolderIsCreated()
    ' The createProject subroutine must create a new folder named Tests
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:="TestProject", displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\TestProject\Tests", vbDirectory) <> "", "The Tests folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub
Public Sub TestConfProdFolderIsCreated()
    ' The createProject subroutine must create a new folder named ConfProd
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:="TestProject", displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\TestProject\source\ConfProd", vbDirectory) <> "", "source\ConfProd folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub
Public Sub TestConfTestFolderIsCreated()
    ' The createProject subroutine must create a new folder named ConfProd
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:="TestProject", displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\TestProject\source\ConfTest", vbDirectory) <> "", "The source\ConfTest folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub
Public Sub TestVbaUnitFolderIsCreated()
    ' The createProject subroutine must create a new folder named ConfProd
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:="TestProject", displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\TestProject\source\VbaUnit", vbDirectory) <> "", "The source\VbaUnit folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
End Sub
Public Sub Testxlafilecreated()
    ' The function how create excel project
    Dim name  As String 'for file .xla
    name = "TestProject" 'initialize it
    Dim returnValue As Long
    Dim returnValue2 As Long
    
    returnValue = vtkCreateProject(path:=vtkTestPath, name:="TestProject", displayError:=False)
    
    returnValue2 = createxlsfile(path:=vtkTestPath, name:="TestProject", displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & name & "\" & "Project" & "\" & name & ".xls", vbDirectory) <> "", "the work book must be created"
   
    mAssert.Equals returnValue, 0, "the workbook creation must be succeded"
End Sub
Public Sub Testvbaunitexported()
    Dim returnValue As Long
    Dim returnValue2 As Long
    Dim returnValue3 As Long
    Dim name  As String 'for file .xla
    name = "TestProject" 'initialize it
    returnValue = vtkCreateProject(path:=vtkTestPath, name:="TestProject", displayError:=False)
    returnValue2 = createxlsfile(path:=vtkTestPath, name:="TestProject", displayError:=False)
    returnValue3 = exportvbaunitwithproject(name:="TestProject")
    mAssert.Should returnValue3 = 18, "the exported module number must be 18"
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
    ITest_Suite.AddTest ITest_Manager.ClassName, "Testxlafilecreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Testvbaunitexported"
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
        Case "Testxlafilecreated": Testxlafilecreated
        Case "Testvbaunitexported": Testvbaunitexported
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

