VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkCreateProjectTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkCreateProjectTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the Create Project main function
'
' Copyright 2013 Skwal-Soft (http://skwalsoft.com)
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
'---------------------------------------------------------------------------------------

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
End Sub

Private Sub ITestCase_TearDown()
    vtkResetConfigurationManagers             ' reset all configuration managers
    'Make sure to be out of the folder to clean
    Dir (vtkTestPath)
    'to not display messagebox how ask to save project
    Application.DisplayAlerts = False
    On Error Resume Next
    'Close Created WorkBook
    Workbooks(vtkTestProjectName & ".xlsm").Close
    Workbooks(vtkTestProjectName & "_DEV" & ".xlsm").Close
    'Delete Created WorkBook
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Project\" & vtkTestProjectName & "_DEV.xlsm"
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Delivery\" & vtkTestProjectName & ".xlsm"
'   Delete Folder Tree
    vtkCleanFolder vtkTestPath
    Application.DisplayAlerts = True
End Sub

Private Sub OpenDeliveryProject()
    Dim path As String
    Dim fso As New FileSystemObject
    path = fso.GetParentFolderName(ActiveWorkbook.path) & "\" & vtkConfigurationManagerForProject(vtkTestProjectName).getConfigurationPath(vtkTestProjectName)
    Workbooks.Open fileName:=path
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
    mAssert.Equals returnvalue, VTK_OK, "The project creation must succeed"
End Sub

Public Sub TestDevXlsmFileCreated()
'test that xlsm file is created in the right path
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & "_DEV.xlsm", vbDirectory) <> "", "the workbook must be created"
End Sub

Public Sub TestDevXlsmFileIsNamedcorrectly()
'test that xlsm is correctly renamed
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Equals Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.name, vtkTestProjectName & "_DEV", "the workbook name must be like " & vtkTestProjectName & "_DEV"
End Sub

Public Sub TestDelivXlsmFileCreated()
'test that xlsm file is created in the right path
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\" & "Delivery" & "\" & vtkTestProjectName & ".xlsm", vbDirectory) <> "", "the workbook must be created"
End Sub

Public Sub TestDevProjectIsNamedcorrectly()
'test that xlsm is correctly renamed
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.name = vtkTestProjectName & "_DEV", "the workbook name must be like " & vtkTestProjectName & "_DEV"
End Sub

Public Sub TestDelivProjectIsNamedcorrectly()
'test that xlsm is correctly renamed
    Dim returnvalue As Long
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    OpenDeliveryProject
    mAssert.Equals Workbooks(vtkTestProjectName & ".xlsm").VBProject.name, vtkTestProjectName, "the workbook name must be like " & vtkTestProjectName & "_DEV"
End Sub

Public Sub TestReferencesOfDevWorkbookWasActivated()
 Dim returnvalue As Long
 Dim i As Integer
 Dim j As Integer
 j = 0
'this function will count extention number before executing function ,
'and extentions number must be equal to(initial extention number + added extention number )
returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    OpenDeliveryProject
 
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
    OpenDeliveryProject
    
    For i = 1 To Workbooks(vtkTestProjectName & ".xlsm").VBProject.References.Count
        If ((Workbooks(vtkTestProjectName & ".xlsm").VBProject.References(i).GUID) = "{420B2830-E718-11CF-893D-00A0C9054228}") Or (Workbooks(vtkTestProjectName & ".xlsm").VBProject.References(i).GUID = "{0002E157-0000-0000-C000-000000000046}") Then
        j = j + 1
        End If
    Next
mAssert.Equals j, 2, "the references number must be equal to" & i + 2
End Sub

Public Sub Test_VBAUnitModulesExportation()
    '   Verify that VBA modules are exported from VBAToolkit to the new project
    Dim returnvalue As Long, fso As New FileSystemObject
    
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    
    mAssert.Should fso.FileExists(vtkTestPath & "\" & vtkTestProjectName & "\" & "Source\VbaUnit\" & "VbaUnitMain.bas"), "VBAUnitMain module have to be exported"
End Sub

Public Sub Test_VBAUnitModulesImportation()
    '   Verify that VBA modules are imported in the new project
    Dim returnvalue As Long, fso As New FileSystemObject
    
    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    
   On Error GoTo Test_VBAUnitModulesImportation_Error
    mAssert.Should Not Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.VBComponents("VBAUnitMain") Is Nothing, "VBAUnitMain module have to be imported"
    mAssert.Should Not Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.VBComponents("TestClassLister") Is Nothing, "TestClassLister module have to be imported"
    mAssert.Should Not Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.VBComponents("TestRunner") Is Nothing, "TestRunner module have to be imported"
    mAssert.Should Not Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.VBComponents("IRunManager") Is Nothing, "IRunManager module have to be imported"
    mAssert.Should Not Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.VBComponents("TestResult") Is Nothing, "TestResult module have to be imported"
   On Error GoTo 0
   Exit Sub

Test_VBAUnitModulesImportation_Error:
    mAssert.Should False, "No error have to be raised during project creation (Err=" & err.Number & ": " & err.Description & ")"
End Sub

Public Sub Test_vtkCreateProject_correctPathGitFolderIsCreated()
    Dim fso As New FileSystemObject
    
    On Error GoTo Test_vtkCreateProject_correctPathGitFolderIsCreated_Error

    vtkCreateProject vtkTestPath, vtkTestProjectName, displayError:=False
    
    mAssert.Should fso.folderExists(vtkTestPath & "\" & vtkTestProjectName & "\.git"), ".git folder should be created."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateProject_correctPathGitFolderIsCreated_Error:
    mAssert.Should False, "Error " & err.Number & " (" & err.Description & ") in procedure Test_vtkCreateProject_gitFolderIsCreated of Class Module vtkCreateProjectTester"
    Resume Next
End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateProjectDoesNothingWhenPathDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateProjectAlreadyExisting"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestMainFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDevXlsmFileCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDevXlsmFileIsNamedcorrectly"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDelivXlsmFileCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDevProjectIsNamedcorrectly"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDelivProjectIsNamedcorrectly"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferencesOfDevWorkbookWasActivated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferencesOfDeliveryWorkbookWasActivated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_VBAUnitModulesExportation"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_VBAUnitModulesImportation"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateProject_correctPathGitFolderIsCreated"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestCreateProjectDoesNothingWhenPathDoesntExist": TestCreateProjectDoesNothingWhenPathDoesntExist
        Case "TestCreateProjectAlreadyExisting": TestCreateProjectAlreadyExisting
        Case "TestMainFolderIsCreated": TestMainFolderIsCreated
        Case "TestDevXlsmFileCreated": TestDevXlsmFileCreated
        Case "TestDevXlsmFileIsNamedcorrectly": TestDevXlsmFileIsNamedcorrectly
        Case "TestDelivXlsmFileCreated": TestDelivXlsmFileCreated
        Case "TestDevProjectIsNamedcorrectly": TestDevProjectIsNamedcorrectly
        Case "TestDelivProjectIsNamedcorrectly": TestDelivProjectIsNamedcorrectly
        Case "TestReferencesOfDevWorkbookWasActivated": TestReferencesOfDevWorkbookWasActivated
        Case "TestReferencesOfDeliveryWorkbookWasActivated": TestReferencesOfDeliveryWorkbookWasActivated
        Case "Test_VBAUnitModulesExportation": Test_VBAUnitModulesExportation
        Case "Test_VBAUnitModulesImportation": Test_VBAUnitModulesImportation
        Case "Test_vtkCreateProject_correctPathGitFolderIsCreated": Test_vtkCreateProject_correctPathGitFolderIsCreated
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

