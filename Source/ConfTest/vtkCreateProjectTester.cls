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

Private Const vtkTestProjectName = "TestProject"
Private Const expectedReferenceCount = 9

Private mManager As TestCaseManager
Private mAssert As IAssert

Private fso As New FileSystemObject

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
    vtkCleanFolder VBAToolKit.vtkTestPath
    vtkResetConfigurationManagers             ' reset all configuration managers
    vtkResetProjects
End Sub

Private Sub ITestCase_TearDown()
    vtkCleanFolder VBAToolKit.vtkTestPath
    'VBAToolKit.resetTestFolder 'Throws me a "Error 5006 : permission denied" at third test
    vtkResetConfigurationManagers             ' reset all configuration managers
    vtkResetProjects
End Sub

Public Sub TestCreateProjectDoesNothingWhenPathDoesntExist()
    ' The createProject must do nothing when path doesn't exist
    Dim InexistentPath As String
    Dim returnValue As Long
    On Error GoTo TestCreateProjectDoesNothingWhenPathDoesntExist_Error

    InexistentPath = VBAToolKit.vtkTestPath & "\InexistentFolder"
    returnValue = vtkCreateProject(path:=InexistentPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(InexistentPath & "\" & vtkTestProjectName, vbDirectory) = "", "The createProject must do nothing when path doesn't exist"
    mAssert.Equals returnValue, 76, "The project creation must fail"

    On Error GoTo 0
    Exit Sub

TestCreateProjectDoesNothingWhenPathDoesntExist_Error:
    Err.Source = "Sub TestCreateProjectDoesNothingWhenPathDoesntExist in module vtkCreateProjectTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub TestCreateProjectAlreadyExisting()
    ' The createProject must does nothing when project already exists
    Dim returnValue As Long
    On Error GoTo TestCreateProjectAlreadyExisting_Error

    MkDir VBAToolKit.vtkTestPath & "\" & vtkTestProjectName 'must respect order mkdir before call vtkcreateproject
    returnValue = vtkCreateProject(path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & vtkTestProjectName), "The existing project main folder must exist"
    mAssert.Equals returnValue, 75, "The project creation must fail"

    On Error GoTo 0
    Exit Sub

TestCreateProjectAlreadyExisting_Error:
    Err.Source = "Sub TestCreateProjectAlreadyExisting in module vtkCreateProjectTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub TestMainFolderIsCreated()
    ' The createProject subroutine must create a new folder named like the project in the path
    Dim returnValue As Long
    On Error GoTo TestMainFolderIsCreated_Error

    returnValue = vtkCreateProject(path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & vtkTestProjectName), "The project main folder must exist"
    mAssert.Equals returnValue, VTK_OK, "The project creation must succeed"

    On Error GoTo 0
    Exit Sub

TestMainFolderIsCreated_Error:
    Err.Source = "Sub TestMainFolderIsCreated in module vtkCreateProjectTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub TestDevXmlFileCreated()
'test that xml file is created in the right path
    Dim returnValue As Long
    On Error GoTo TestDevXmlFileCreated_Error

    returnValue = vtkCreateProject(path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & ".xml"), "the xml sheet must be created"

    On Error GoTo 0
    Exit Sub

TestDevXmlFileCreated_Error:
    Err.Source = "Sub TestDevXmlFileCreated in module vtkCreateProjectTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub TestProjectIsNamedcorrectly()
'test that the project is named correctly in the configuration sheet
    On Error GoTo TestProjectIsNamedcorrectly_Error

    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    
    Dim dom As New MSXML2.DOMDocument
    dom.Load VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & ".xml"
    
    mAssert.Equals dom.SelectSingleNode("/vtkConf/info/projectName[0]").Text, vtkTestProjectName, "the project name in the sheet must be like " & vtkTestProjectName

    On Error GoTo 0
    Exit Sub

TestProjectIsNamedcorrectly_Error:
    Err.Source = "Sub TestProjectIsNamedcorrectly in module vtkCreateProjectTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub TestDevProjectIsNamedcorrectly()
'test that the dev configuration is named correctly in the configuration sheet
    On Error GoTo TestDevProjectIsNamedcorrectly_Error
    
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    
    Dim dom As New MSXML2.DOMDocument
    dom.Load VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & ".xml"
    
    mAssert.Equals dom.SelectSingleNode("/vtkConf/configuration[@cID=""c01""]/name[0]").Text, vtkTestProjectName & "_DEV", "the dev configuration must be called " & vtkTestProjectName & "_DEV"

    On Error GoTo 0
    Exit Sub

TestDevProjectIsNamedcorrectly_Error:
    Err.Source = "Sub TestDevProjectIsNamedcorrectly in module vtkCreateProjectTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub TestDelivProjectIsNamedcorrectly()
'test that the delivery configuration is named correctly in the configuration sheet
    On Error GoTo TestDelivProjectIsNamedcorrectly_Error

    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    
    Dim dom As New MSXML2.DOMDocument
    dom.Load VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & ".xml"
    
    mAssert.Equals dom.SelectSingleNode("/vtkConf/configuration[@cID=""c02""]/name[0]").Text, vtkTestProjectName, "the delivery configuration must be called " & vtkTestProjectName
 
    On Error GoTo 0
    Exit Sub

TestDelivProjectIsNamedcorrectly_Error:
    Err.Source = "Sub TestDelivProjectIsNamedcorrectly in module vtkCreateProjectTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkCreateProject_correctPathGitFolderIsCreated()
' test that the .git folder is created
    Dim fso As New FileSystemObject

    On Error GoTo Test_vtkCreateProject_correctPathGitFolderIsCreated_Error

    vtkCreateProject VBAToolKit.vtkTestPath, vtkTestProjectName, displayError:=False

    mAssert.Should fso.folderExists(VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\.git"), ".git folder should be created."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateProject_correctPathGitFolderIsCreated_Error:
    mAssert.Should False, "Error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkCreateProject_gitFolderIsCreated of Class Module vtkCreateProjectTester"
    Resume Next
End Sub


Public Sub TestDefaultReferencesArePresent()
' test that the default VBAToolKit references are present in the sheet
    On Error GoTo TestDefaultReferencesArePresent_Error

    vtkCreateProject path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False
    
    Dim dom As New MSXML2.DOMDocument
    dom.Load VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & ".xml"
    
    Dim tmpRef As VBIDE.Reference
    For Each tmpRef In listOfDefaultReferences
        mAssert.Equals dom.SelectNodes("/vtkConf/reference[name=""" & tmpRef.name & """]").Length, 1, "The reference " & tmpRef.name & " should be present in the sheet."
    Next

    On Error GoTo 0
    Exit Sub

TestDefaultReferencesArePresent_Error:
    Err.Source = "Sub TestDefaultReferencesArePresent in module vtkCreateProjectTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub TestToSelfReferenceIsPresentAndActivatedByDefaultForDEV()
' test that the DEV configuration has a reference to the workbook that created the project
    On Error GoTo TestToSelfReferenceIsPresentAndActivatedByDefaultForDEV_Error
    
    vtkCreateProject path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False
    
    Dim dom As New MSXML2.DOMDocument
    dom.Load VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & ".xml"
    
    mAssert.Equals dom.SelectNodes("/vtkConf/reference[@confIDs=""c01""][name=""" & ThisWorkbook.VBProject.name & """]").Length, 1, "The reference to the current workbook must be present."

    On Error GoTo 0
    Exit Sub

TestToSelfReferenceIsPresentAndActivatedByDefaultForDEV_Error:
    Err.Source = "Sub TestToSelfReferenceIsPresentAndActivatedByDefaultForDEV in module vtkCreateProjectTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

'Public Sub TestReferencesOfDevConfigurationWasActivated()
'' test that the references of the DEV configuration are activated
'    On Error GoTo TestReferencesOfDevConfigurationWasActivated_Error
'
'    vtkCreateProject path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False
'
'    Dim dom As New MSXML2.DOMDocument
'    dom.Load VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & ".xml"
'
'    Dim confId As String
'    confId = dom.SelectSingleNode("/vtkConf/configuration[name=""" & vtkTestProjectName & "_DEV""]/@cID").Text
'
'    Dim tmpRef As VBIDE.Reference
'    For Each tmpRef In listOfDefaultReferences
'        mAssert.Equals dom.SelectNodes("/vtkConf/reference[name=""" & tmpRef.name & """]/@confId='" & confId & "'").Length, 1, "The reference " & tmpRef.name & " should be present in the sheet."
'    Next
'
'    On Error GoTo 0
'    Exit Sub
'
'TestReferencesOfDevConfigurationWasActivated_Error:
'    Err.Source = "Sub TestReferencesOfDevConfigurationWasActivated in module vtkCreateProjectTester"
'    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
'    Exit Sub
'
'
'End Sub
'''
'''
'''Public Sub TestReferencesOfDeliveryWorkbookWasActivated()
''' Dim returnValue As Long
''' Dim i As Integer
''' Dim j As Integer
''' j = 0
''''this function will count extention number before executing function ,
''''and extentions number must be equal to(initial extention number + added extention number )
'''returnValue = vtkCreateProject(path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'''    OpenDeliveryProject
'''
'''    For i = 1 To Workbooks(vtkTestProjectName & ".xlsm").VBProject.References.count
'''        If ((Workbooks(vtkTestProjectName & ".xlsm").VBProject.References(i).GUID) = "{420B2830-E718-11CF-893D-00A0C9054228}") Or (Workbooks(vtkTestProjectName & ".xlsm").VBProject.References(i).GUID = "{0002E157-0000-0000-C000-000000000046}") Then
'''        j = j + 1
'''        End If
'''    Next
'''mAssert.Equals j, 2, "the searched reference number must be equal to " & 2
'''mAssert.Equals Workbooks(vtkTestProjectName & ".xlsm").VBProject.References.count, expectedReferenceCount, "the total references number must be equal to " & expectedReferenceCount
'''End Sub
'''
'''Public Sub Test_VBAUnitModulesExportation()
'''    '   Verify that VBA modules are exported from VBAToolkit to the new project
'''    Dim returnValue As Long, fso As New FileSystemObject
'''
'''    returnValue = vtkCreateProject(path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'''
'''    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\" & "Source\VbaUnit\" & "VbaUnitMain.bas"), "VBAUnitMain module have to be exported"
'''End Sub
'''
'''Public Sub Test_VBAUnitModulesImportation()
'''    '   Verify that VBA modules are imported in the new project
'''    Dim returnValue As Long, fso As New FileSystemObject
'''
'''    returnValue = vtkCreateProject(path:=VBAToolKit.vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'''
'''
'''   On Error GoTo Test_VBAUnitModulesImportation_Error
'''    mAssert.Should Not Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.VBComponents("VBAUnitMain") Is Nothing, "VBAUnitMain module have to be imported"
'''    mAssert.Should Not Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.VBComponents("TestClassLister") Is Nothing, "TestClassLister module have to be imported"
'''    mAssert.Should Not Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.VBComponents("TestRunner") Is Nothing, "TestRunner module have to be imported"
'''    mAssert.Should Not Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.VBComponents("IRunManager") Is Nothing, "IRunManager module have to be imported"
'''    mAssert.Should Not Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.VBComponents("TestResult") Is Nothing, "TestResult module have to be imported"
'''   On Error GoTo 0
'''   Exit Sub
'''
'''Test_VBAUnitModulesImportation_Error:
'''    mAssert.Should False, "No error have to be raised during project creation (Err=" & Err.Number & ": " & Err.Description & ")"
'''End Sub
'''
'''
'''Public Sub Test_vtkCreateProject_handlerAddedInDEVWorkbook_modulesAreExported()
'''' creating a project, modifying an existing module in the DEV workbook, save => module must be exported
'''
'''    Dim retVal As Integer
'''    Dim oldSize1 As Integer
'''    Dim oldSize2 As Integer
'''    Dim newSize1 As Integer
'''    Dim newSize2 As Integer
'''    Dim newlyCreatedWorkbook As Workbook
'''    Dim fso As New FileSystemObject
'''
'''    On Error GoTo Test_vtkAddBeforeSaveHandlerInDEVWorkbook_justCreatedWorkbook_Error
'''
'''    vtkCreateProject VBAToolKit.vtkTestPath, vtkTestProjectName
'''
'''    Set newlyCreatedWorkbook = Workbooks(vtkTestProjectName & "_DEV.xlsm")
'''
'''    ' Get the size of the exported modules before the save
'''    oldSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit\VbaUnitMain.bas").Size
'''    oldSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit\Assert.cls").Size
'''
'''    ' Modify a module
'''    insertDummyProcedureInCodeModule newlyCreatedWorkbook.VBProject.VBComponents("VbaUnitMain").CodeModule, dummyInt:=1
'''
'''    ' Save the workbook
'''    newlyCreatedWorkbook.Save
'''
'''    ' Get the new size of the exported modules
'''    newSize1 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit\VbaUnitMain.bas").Size
'''    newSize2 = fso.GetFile(VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit\Assert.cls").Size
'''
'''    mAssert.Should oldSize1 < newSize1, "VbaUnitMain.bas must be bigger after the save."
'''    mAssert.Should oldSize2 = newSize2, "Assert.cls must be of equal size after the save."
'''
'''    On Error GoTo 0
'''    Exit Sub
'''
'''Test_vtkAddBeforeSaveHandlerInDEVWorkbook_justCreatedWorkbook_Error:
'''    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkAddBeforeSaveHandlerInDEVWorkbook_justCreatedWorkbook of Class Module vtkBeforeSaveHandlerTester"
'''    ITestCase_TearDown
'''    Exit Sub
'''
'''End Sub
'''
'''Public Sub Test_vtkCreateProject_handlerAddedInDEVWorkbook_xmlIsCreated()
'''' creating a project : the xml export of its configuration should exist as the DEV workbook is saved in vtkCreateProject
'''
'''    Dim newlyCreatedWorkbook As Workbook
'''    Dim fso As New FileSystemObject
'''
'''    On Error GoTo Test_vtkCreateProject_handlerAddedInDEVWorkbook_xmlIsCreated_Error
'''
'''    vtkCreateProject VBAToolKit.vtkTestPath, vtkTestProjectName
'''
'''    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & vtkTestProjectName & "\Project\" & vtkTestProjectName & ".xml") = True, _
'''                   "The XML file should exist at this point."
'''
'''    On Error GoTo 0
'''    Exit Sub
'''
'''Test_vtkCreateProject_handlerAddedInDEVWorkbook_xmlIsCreated_Error:
'''    Err.Source = "Test_vtkCreateProject_handlerAddedInDEVWorkbook_xmlIsCreated of module vtkCreateProjectTester"
'''    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
'''    ITestCase_TearDown
'''    Exit Sub
'''End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateProjectDoesNothingWhenPathDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateProjectAlreadyExisting"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestMainFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDevXmlFileCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectIsNamedcorrectly"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDevProjectIsNamedcorrectly"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDelivProjectIsNamedcorrectly"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateProject_correctPathGitFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDefaultReferencesArePresent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestToSelfReferenceIsPresentAndActivatedByDefaultForDEV"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestCreateProjectDoesNothingWhenPathDoesntExist": TestCreateProjectDoesNothingWhenPathDoesntExist
        Case "TestCreateProjectAlreadyExisting": TestCreateProjectAlreadyExisting
        Case "TestMainFolderIsCreated": TestMainFolderIsCreated
        Case "TestDevXmlFileCreated": TestDevXmlFileCreated
        Case "TestProjectIsNamedcorrectly": TestProjectIsNamedcorrectly
        Case "TestDevProjectIsNamedcorrectly": TestDevProjectIsNamedcorrectly
        Case "TestDelivProjectIsNamedcorrectly": TestDelivProjectIsNamedcorrectly
        Case "Test_vtkCreateProject_correctPathGitFolderIsCreated": Test_vtkCreateProject_correctPathGitFolderIsCreated
        Case "TestDefaultReferencesArePresent": TestDefaultReferencesArePresent
        Case "TestToSelfReferenceIsPresentAndActivatedByDefaultForDEV": TestToSelfReferenceIsPresentAndActivatedByDefaultForDEV
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

