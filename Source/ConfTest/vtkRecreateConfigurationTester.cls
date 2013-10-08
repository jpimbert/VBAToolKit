VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkRecreateConfigurationTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkRecreateConfigurationTester
' Author    : Lucas Vitorino
' Purpose   : Test the vtkRecreateConfiguration function
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

Private fso As New FileSystemObject

Private Wb As Workbook
Private recreatedWb As Workbook
Private testProject As New vtkProject
Private normalConf As vtkConfiguration
Private emptyConf As vtkConfiguration
Private wrongPathConf As vtkConfiguration
Private xlsConf As vtkConfiguration
Private xlamConf As vtkConfiguration
Private xlaConf As vtkConfiguration
Private testedConf As vtkConfiguration

Private cm As vtkConfigurationManager

Private Const testProjectName = "TestProjectForRecreateConf"
Private Const normalConfName = "TestProject_normalConf"
Private Const emptyConfName = "TestProject_emptyConf"
Private Const wrongPathConfName = "TestProject_wrongPathConf"
Private Const xlsConfName = "TestProject_xlsPathConf"
Private Const xlamConfName = "TestProject_xlamConf"
Private Const xlaConfName = "TestProject_xlaConf"

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
    
    Set testProject = vtkProjectForName(testProjectName)
    
    Set Wb = getTestFileFromTemplate(testProject.workbookDEVName, openExcel:=True)

    Set cm = vtkConfigurationManagerForProject(testProjectName)
    Set normalConf = cm.configurations(normalConfName)
    Set emptyConf = cm.configurations(emptyConfName)
    Set wrongPathConf = cm.configurations(wrongPathConfName)
    Set xlsConf = cm.configurations(xlsConfName)
    Set xlamConf = cm.configurations(xlamConfName)
    Set xlaConf = cm.configurations(xlaConfName)
    
End Sub

Private Sub ITestCase_TearDown()
    On Error Resume Next
    vtkCloseAndKillWorkbook Wb
    vtkCloseAndKillWorkbook recreatedWb
    VBAToolKit.resetTestFolder
    vtkResetProjects
    vtkResetConfigurationManagers
End Sub

' TESTS TO BE WRITTEN
'
' normal configuration, normal path
'   - workbook is created in the good path
'   - all modules of the configuration are present
'   - references are activated ( will probably change in the near future as reference management in configurations will change)
'   - name of the vbproject is correct : same name as the recreated configuration
'   - filetype is consistent with the extension
'       - xlsm
'       - xlam
'       - xls
'       - xla
'
' configuration particular cases
'   - configuration is empty
'   - different kind of modules (noraml, class, form)
'
' particular cases
'   - workbook with the same name is already open
'   - workbook with the same name already exists in the same path
'   - source files have not been yet exported
'

Public Sub Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath()
' workbook is created in the good path

    On Error GoTo Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_Error
    
    Set testedConf = normalConf
    
    vtkRecreateConfiguration testProject.projectName, testedConf.name

    mAssert.Should fso.FileExists(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path), _
                   "The configuration should have been created in the designated path."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_Error:
    Err.Source = "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfiguration_allModulesArePresent()
' all modules of the configuration are present

    Dim error As Integer: error = 0

    On Error GoTo Test_vtkRecreateConfiguration_allModulesArePresent_Error
    
    Set testedConf = normalConf
    
    ' It is mandatory to create the source files before recreating a configuration
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration testProject.projectName, testedConf.name
    
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)
    
    Dim mo As vtkModule
    Dim vbCmp As VBComponent
    For Each mo In testedConf.modules
        ' Just a dummy access attempt : will throw error 9 (subscript out of range) if the module doesn't exist
        Set vbCmp = recreatedWb.VBProject.VBComponents(mo.name)
    Next
    
    mAssert.Equals error, 0, "All modules don't exist in the recreated project."
    
    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_allModulesArePresent_Error:
        If Err.Number = 9 Then
            error = Err.Number
            Resume Next
        Else
            Err.Source = "Test_vtkRecreateConfiguration_allModulesArePresent of module vtkRecreateConfigurationTester"
            mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
            ITestCase_TearDown
            Exit Sub
        End If
End Sub

Public Sub Test_vtkRecreateConfiguration_referencesAreActivated()
' references are activated
' will probably change in the near future as reference management in configurations will change

    On Error GoTo Test_vtkRecreateConfiguration_referencesAreActivated_Error

    Set testedConf = normalConf

    Dim numRef As Integer
    Dim ref As Reference
    Dim i As Integer
    Dim refIsActive As Boolean
    Dim guidArray() As String
    
    ' Creation of the new workbook
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)
    
    ' Listing expected references
    numRef = 4
    ReDim guidArray(1 To numRef) As String
    guidArray(1) = "{420B2830-E718-11CF-893D-00A0C9054228}" 'Scripting : Microsoft scripting runtime
    guidArray(2) = "{0002E157-0000-0000-C000-000000000046}" ' VBIDE : Microsoft visual basic for applications extensibility 5.3
    guidArray(3) = "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}" ' Shell32 : Microsoft Shell Controls and Automation
    guidArray(4) = "{F5078F18-C551-11D3-89B9-0000F81FE221}" ' MSXML2 : Microsoft XML V5.0
    
    ' For each GUID in the array, check if it is in the project references
    For i = 1 To numRef
    
        refIsActive = False
        For Each ref In recreatedWb.VBProject.References
            If ref.GUID = guidArray(i) Then refIsActive = True
        Next
        
        mAssert.Should refIsActive, "The reference with GUID = " & guidArray(i) & " is not activated : it should."
    Next
    
    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_referencesAreActivated_Error:
    Err.Source = "Test_vtkRecreateConfiguration_referencesAreActivated of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect()
' name of the vbproject is correct : same name as the recreated configuration

    On Error GoTo Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect_Error

    Set testedConf = normalConf

    ' Creation of the new workbook
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)
    
    mAssert.Equals recreatedWb.VBProject.name, testedConf.name, _
                   "The created project should have the same name as the recreated configuration."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect_Error:
    Err.Source = "Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub


Public Sub Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLSM()

    On Error GoTo Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLSM_Error

    Set testedConf = normalConf

    ' Creation of the new workbook
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)
    
    mAssert.Equals fso.GetFile("VBAToolKit.vtkPathOfCurrentProject(wb) & " \ " & testedConf.path").Type = _
                   xlOpenXMLWorkbookMacroEnabled, _
                   "File format is wrong."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLSM_Error:
    Err.Source = "Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLSM of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLAM()

    On Error GoTo Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLAM_Error
    
    Set testedConf = xlamConf
    
    ' Creation of the new workbook
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)
    
    mAssert.Equals fso.GetFile("VBAToolKit.vtkPathOfCurrentProject(wb) & " \ " & testedConf.path").Type = _
                   xlOpenXMLAddIn, _
                   "File format is wrong."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLAM_Error:
    Err.Source = "Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLAM of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLS()

    On Error GoTo Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLS_Error

    Set testedConf = xlsConf

    ' Creation of the new workbook
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)
    
    mAssert.Equals fso.GetFile("VBAToolKit.vtkPathOfCurrentProject(wb) & " \ " & testedConf.path").Type = _
                   xlExcel8, _
                   "File format is wrong."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLS_Error:
    Err.Source = "Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLS of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLA()

    On Error GoTo Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLA_Error

    Set testedConf = xlaConf

    ' Creation of the new workbook
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)
    
    mAssert.Equals fso.GetFile("VBAToolKit.vtkPathOfCurrentProject(wb) & " \ " & testedConf.path").Type = _
                   xlAddIn8, _
                   "File format is wrong."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLA_Error:
    Err.Source = "Test_vtkRecreateConfiguration_fileTypeIsConsistentWithExtension_XLA of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfiguration_emptyConfiguration()

End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_allModulesArePresent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_referencesAreActivated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath
        Case "Test_vtkRecreateConfiguration_allModulesArePresent": Test_vtkRecreateConfiguration_allModulesArePresent
        Case "Test_vtkRecreateConfiguration_referencesAreActivated": Test_vtkRecreateConfiguration_referencesAreActivated
        Case "Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect": Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

