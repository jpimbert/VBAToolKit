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
End Sub


Private Sub ITestCase_TearDown()

    'Make sure to be out of the folder to clean
    Dir (vtkTestPath)
    'to not display messagebox how ask to save project
    Application.DisplayAlerts = False
    On Error Resume Next
    'Close Created WorkBook
    Workbooks(vtkTestProjectName & ".xlsm").Close
    Workbooks(vtkTestProjectName & "_DEV" & ".xlsm").Close
    'Delete Created WorkBook
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Project\" & vtkTestProjectName & ".xlsm"
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Delivery\" & vtkTestProjectName & "_Delivery" & ".xlsm"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\ConfProd"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\ConfTest"
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit\*"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\GitLog"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Tests"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Delivery"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Project"
    RmDir vtkTestPath & "\" & vtkTestProjectName
    Application.DisplayAlerts = True
End Sub

Public Sub TestCreateProjectDoesNothingWhenPathDoesntExist()
    ' The createProject must does nothing when path doesn't exist
    Dim InexistentPath As String
    Dim returnvalue As Long
    InexistentPath = vtkTestPath & "\InexistentFolder"
    returnvalue = vtkCreateProject(path:=InexistentPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(InexistentPath & "\" & vtkTestProjectName, vbDirectory) = "", "The project main folder must exist"
    mAssert.Equals returnvalue, 76, "The project creation must fail"
End Sub

Public Sub TestCreateProjectAlreadyExisting()
    ' The createProject must does nothing when project already exists
    Dim returnvalue As Long
    MkDir vtkTestPath & "\" & vtkTestProjectName 'must respect order mkdir before call vtkcreateproject
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName, vbDirectory) <> "", "The existing project main folder must exist"
    mAssert.Equals returnvalue, 75, "The project creation must fail"
End Sub

Public Sub TestMainFolderIsCreated()
    ' The createProject subroutine must create a new folder named like the project in the path
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName, vbDirectory) <> "", "The project main folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestProjectFolderIsCreated()
    ' The createProject subroutine must create a new folder named Project
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Project", vbDirectory) <> "", "The Project folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestSourceFolderIsCreated()
    ' The createProject subroutine must create a new folder named Source
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Source", vbDirectory) <> "", "The Source folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestTestsFolderIsCreated()
    ' The createProject subroutine must create a new folder named Tests
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Tests", vbDirectory) <> "", "The Tests folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub
Public Sub TestConfProdFolderIsCreated()
    ' The createProject subroutine must create a new folder named ConfProd
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\ConfProd", vbDirectory) <> "", "source\ConfProd folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub
Public Sub TestConfTestFolderIsCreated()
    ' The createProject subroutine must create a new folder named ConfProd
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\ConfTest", vbDirectory) <> "", "The source\ConfTest folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub
Public Sub TestVbaUnitFolderIsCreated()
    ' The createProject subroutine must create a new folder named ConfProd
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\VbaUnit", vbDirectory) <> "", "The source\VbaUnit folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub
Public Sub TestGitLogIsCreated()
    ' The createProject subroutine must create a new folder named GitLog
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\GitLog", vbDirectory) <> "", "The ..\GitLog folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub
Public Sub TestDeliveryFolderIsCreated()
    ' The createProject subroutine must create a new folder named Delivery
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Delivery", vbDirectory) <> "", "The \Delivery folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub
Public Sub TestDevXlsmFileCreated()
'test that xlsm file is created in the right path
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & ".xlsm", vbDirectory) <> "", "the workbook must be created"
End Sub
Public Sub TestDevXlsmFileIsNamedcorrectly()
'test that xlsm is correctly renamed
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.name = vtkTestProjectName & "_DEV", "the workbook name must be like " & vtkTestProjectName & "_DEV"
End Sub
Public Sub TestDelivXlsmFileCreated()
'test that xlsm file is created in the right path
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\" & "Delivery" & "\" & vtkTestProjectName & "_Delivery" & ".xlsm", vbDirectory) <> "", "the workbook must be created"
End Sub
Public Sub TestDelivXlsmFileIsNamedcorrectly()
'test that xlsm is correctly renamed
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.name = vtkTestProjectName & "_DEV", "the workbook name must be like " & vtkTestProjectName & "_DEV"
End Sub
Public Sub TestReferencesOfDevWorkbookWasActivated()
 Dim returnvalue As Long
 Dim i As Integer
 Dim j As Integer
 j = 0
'this function will count extention number before executing function ,
'and extentions number must be equal to(initial extention number + added extention number )
returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
 
    For i = 1 To Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.References.Count
        If ((Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.References(i).GUID) = "{420B2830-E718-11CF-893D-00A0C9054228}") Or (Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.References(i).GUID = "{0002E157-0000-0000-C000-000000000046}") Then
        j = j + 1
        End If
    Next
mAssert.Equals j, 2, "the references number must be equal to" & i + 2
End Sub
Public Sub TestReferencesOfDeliveryWorkbookWasActivated()
 Dim returnvalue As Long
 Dim i As Integer
 Dim j As Integer
 j = 0
'this function will count extention number before executing function ,
'and extentions number must be equal to(initial extention number + added extention number )
returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
 
    For i = 1 To Workbooks(vtkTestProjectName & ".xlsm").VBProject.References.Count
        If ((Workbooks(vtkTestProjectName & ".xlsm").VBProject.References(i).GUID) = "{420B2830-E718-11CF-893D-00A0C9054228}") Or (Workbooks(vtkTestProjectName & "_Delivery" & ".xlsm").VBProject.References(i).GUID = "{0002E157-0000-0000-C000-000000000046}") Then
        j = j + 1
        End If
    Next
mAssert.Equals j, 2, "the references number must be equal to" & i + 2
End Sub
Public Sub TestDevWorkbookPathCorrectlyWritted()
'test that path  of dev workbook is correctly writted on configuration sheet
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & vtkFirstLine - 2), Workbooks(vtkTestProjectName & ".xlsm").FullNameURLEncoded
End Sub
Public Sub TestDevWorkbookNameCorrectlyWritted()
'test that name  of dev workbook is correctly writted  on configuration sheet
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & vtkFirstLine - 3), Workbooks(vtkTestProjectName & ".xlsm").name
End Sub
Public Sub TestDeliveryWorkbookPathCorrectlyWritted()
'test that path  of delivery workbook is correctly writted on configuration sheet
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & vtkFirstLine - 2), Workbooks(vtkTestProjectName & "_Delivery" & ".xlsm").FullNameURLEncoded
End Sub
Public Sub TestDeliveryWorkbookNameCorrectlyWritted()
'test that name  of delivery workbook is correctly writted on configuration sheet
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & vtkFirstLine - 3), Workbooks(vtkTestProjectName & "_Delivery" & ".xlsm").name
End Sub
Public Sub TestDevWorkbookIsActive()
'test that active workbook is the dev one and not delivery
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Equals ActiveWorkbook.name, vtkTestProjectName & ".xlsm", "the dev workbook must be active"
End Sub
Public Sub TestExportVbaUnitModule()
  'test that exported file was not empty , files was correctly named
  'the test will compare "vbaunit" file of vbatoolkit and exported file
  Dim fs As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  
  'call create project function
  returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
  'vbaunit files of vbatoolkit project
  sourceVbaUnitFile = Dir(vtkInstallPath & "\source\VbaUnit\*.*", vbDirectory)
 
 'list of files in the folder
 Do While sourceVbaUnitFile <> ""
    'ignore these two name
    If sourceVbaUnitFile <> "." And sourceVbaUnitFile <> ".." Then
      'test that exported files was not empty
       mAssert.Should IsEmpty(vtkTestPath & "\" & vtkTestProjectName & "\source\VbaUnit\" & sourceVbaUnitFile) = False, sourceVbaUnitFile & "should not be empty"
      'test name, extension, and the existence of the file in the created project
       mAssert.Should fs.FileExists(vtkTestPath & "\" & vtkTestProjectName & "\source\VbaUnit\" & sourceVbaUnitFile) = True, sourceVbaUnitFile & "must exist"
    End If
  'point to the next element
  sourceVbaUnitFile = Dir
 Loop

End Sub
Public Sub TestImportModules()
'test that modules was imported to the test workbook
Dim returnvalue As String
Dim NbrMod1 As Integer

    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    NbrMod1 = Workbooks(vtkTestProjectName & ".xlsm").VBProject.VBComponents.Count
'18 vbaunitmodule + 4 default module (4sheets + 1 workbook )
mAssert.Equals NbrMod1, 18 + 4 + 1, "18 modules must be imported"
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
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGitLogIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDeliveryFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDevXlsmFileCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDevXlsmFileIsNamedcorrectly"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDelivXlsmFileCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDelivXlsmFileIsNamedcorrectly"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferencesOfDevWorkbookWasActivated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferencesOfDeliveryWorkbookWasActivated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDevWorkbookPathCorrectlyWritted"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDevWorkbookNameCorrectlyWritted"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDeliveryWorkbookPathCorrectlyWritted"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDeliveryWorkbookNameCorrectlyWritted"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDevWorkbookIsActive"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestExportVbaUnitModule"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestImportModules"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
'        Case "TestCreateProjectDoesNothingWhenPathDoesntExist": TestCreateProjectDoesNothingWhenPathDoesntExist
'        Case "TestCreateProjectAlreadyExisting": TestCreateProjectAlreadyExisting
'        Case "TestMainFolderIsCreated": TestMainFolderIsCreated
'        Case "TestProjectFolderIsCreated": TestProjectFolderIsCreated
'        Case "TestSourceFolderIsCreated": TestSourceFolderIsCreated
'        Case "TestTestsFolderIsCreated": TestTestsFolderIsCreated
'        Case "TestConfProdFolderIsCreated": TestConfProdFolderIsCreated
'        Case "TestConfTestFolderIsCreated": TestConfTestFolderIsCreated
'        Case "TestVbaUnitFolderIsCreated": TestVbaUnitFolderIsCreated
'        Case "TestGitLogIsCreated": TestGitLogIsCreated
'        Case "TestDeliveryFolderIsCreated": TestDeliveryFolderIsCreated
'        Case "TestDevXlsmFileCreated": TestDevXlsmFileCreated
'        Case "TestDevXlsmFileIsNamedcorrectly": TestDevXlsmFileIsNamedcorrectly
'        Case "TestDelivXlsmFileCreated": TestDelivXlsmFileCreated
'        Case "TestDelivXlsmFileIsNamedcorrectly": TestDelivXlsmFileIsNamedcorrectly
'        Case "TestReferencesOfDevWorkbookWasActivated": TestReferencesOfDevWorkbookWasActivated
        Case "TestReferencesOfDeliveryWorkbookWasActivated": TestReferencesOfDeliveryWorkbookWasActivated
'        Case "TestDevWorkbookPathCorrectlyWritted": TestDevWorkbookPathCorrectlyWritted
'        Case "TestDevWorkbookNameCorrectlyWritted": TestDevWorkbookNameCorrectlyWritted
'        Case "TestDeliveryWorkbookPathCorrectlyWritted": TestDeliveryWorkbookPathCorrectlyWritted
'        Case "TestDeliveryWorkbookNameCorrectlyWritted": TestDeliveryWorkbookNameCorrectlyWritted
'        Case "TestDevWorkbookIsActive": TestDevWorkbookIsActive
'        Case "TestExportVbaUnitModule": TestExportVbaUnitModule
'        Case "TestImportModules": TestImportModules
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

