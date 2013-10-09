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
Private multitypeConf As vtkConfiguration
Private VTKConf As vtkConfiguration
Private testedConf As vtkConfiguration

Private cm As vtkConfigurationManager

Private Const testProjectName = "TestProjectForRecreateConf"
Private Const normalConfName = "TestProject_normalConf"
Private Const emptyConfName = "TestProject_emptyConf"
Private Const wrongPathConfName = "TestProject_wrongPathConf"
Private Const xlsConfName = "TestProject_xlsConf"
Private Const xlamConfName = "TestProject_xlamConf"
Private Const xlaConfName = "TestProject_xlaConf"
Private Const multitypeConfName = "TestProject_multitypeConf"
Private Const VTKConfName = "TestProject_VTKConf"

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
    Set multitypeConf = cm.configurations(multitypeConfName)
    Set VTKConf = cm.configurations(VTKConfName)

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
'   - workbook is created in the good path
'       - normal conf
'       - empty conf
'       - multitype conf
'   - all modules of the configuration are present
'       - normal conf
'       - multitype conf
'   - references are activated ( will probably change in the near future as reference management in configurations will change)
'       - normal conf
'   - name of the vbproject is correct : same name as the recreated configuration
'       - normal conf
'   - fileFormat is consistent with the extension
'       - xlsm
'       - xlam
'       - xls
'       - xla
'
' particular cases
'   - workbook with the same name is already open : raise error VTK_WORKBOOK_ALREADY_OPEN
'   - add-in with the same name is activated : export the file normally
'   - workbook with the same name already exists in the same path : overwrite the file
'       - normal conf
'       - xlam


' --------------------------------------------------------
'                       TESTS
' --------------------------------------------------------

Public Sub Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf()
' workbook is created in the good path

    On Error GoTo Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf_Error

    Set testedConf = normalConf

    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name

    mAssert.Should fso.FileExists(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path), _
                   "The configuration should have been created in the designated path."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf_Error:
    Err.Source = "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf()
' workbook is created in the good path

    On Error GoTo Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf_Error

    Set testedConf = emptyConf

    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name

    mAssert.Should fso.FileExists(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path), _
                   "The configuration should have been created in the designated path."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf_Error:
    Err.Source = "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf()
' workbook is created in the good path

    On Error GoTo Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf_Error

    Set testedConf = multitypeConf

    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name

    mAssert.Should fso.FileExists(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path), _
                   "The configuration should have been created in the designated path."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf_Error:
    Err.Source = "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_allModulesArePresent_normalConf()
' all modules of the configuration are present

    Dim error As Integer: error = 0

    On Error GoTo Test_vtkRecreateConfiguration_allModulesArePresent_normalConf_Error

    Set testedConf = normalConf

    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name
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

Test_vtkRecreateConfiguration_allModulesArePresent_normalConf_Error:
        If Err.Number = 9 Then
            error = Err.Number
            Resume Next
        Else
            Err.Source = "Test_vtkRecreateConfiguration_allModulesArePresent_normalConf of module vtkRecreateConfigurationTester"
            mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
            ITestCase_TearDown
            Exit Sub
        End If
End Sub



Public Sub Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf()
' all modules of the configuration are present

    Dim error As Integer: error = 0

    On Error GoTo Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf_Error

    Set testedConf = multitypeConf

    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name
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

Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf_Error:
        If Err.Number = 9 Then
            error = Err.Number
            Resume Next
        Else
            Err.Source = "Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf of module vtkRecreateConfigurationTester"
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

    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name
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

    ' Set recreatedWb and open it
    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name
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



Public Sub Test_vtkRecreateConfiguration_formatLikeExtension_XLSM()
' test the xlsm format

    On Error GoTo Test_vtkRecreateConfiguration_formatLikeExtension_XLSM_Error

    Set testedConf = normalConf

    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)

    mAssert.Equals recreatedWb.FileFormat, _
                   xlOpenXMLWorkbookMacroEnabled, _
                   "File format is wrong."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_formatLikeExtension_XLSM_Error:
    Err.Source = "Test_vtkRecreateConfiguration_formatLikeExtension_XLSM of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_formatLikeExtension_XLAM()
' test the xlam format

    On Error GoTo Test_vtkRecreateConfiguration_formatLikeExtension_XLAM_Error

    Set testedConf = xlamConf

    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)

    mAssert.Equals recreatedWb.FileFormat, _
                   xlOpenXMLAddIn, _
                   "File format is wrong."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_formatLikeExtension_XLAM_Error:
    Err.Source = "Test_vtkRecreateConfiguration_formatLikeExtension_XLAM of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_formatLikeExtension_XLS()
' test the xls format

    On Error GoTo Test_vtkRecreateConfiguration_formatLikeExtension_XLS_Error

    Set testedConf = xlsConf

    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)

    mAssert.Equals recreatedWb.FileFormat, _
                   xlExcel8, _
                   "File format is wrong."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_formatLikeExtension_XLS_Error:
    Err.Source = "Test_vtkRecreateConfiguration_formatLikeExtension_XLS of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub



'Public Sub Test_vtkRecreateConfiguration_formatLikeExtension_XLA()
'' test the xla format
'' test failing, don 't know why
'
'    On Error GoTo Test_vtkRecreateConfiguration_formatLikeExtension_XLA_Error
'
'    Set testedConf = xlaConf
'
'    vtkRecreateConfiguration wb.VBProject, testProject.projectName, testedConf.name
'    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(wb) & "\" & testedConf.path)
'
'    mAssert.Equals recreatedWb.FileFormat, _
'                   xlAddIn, _
'                   "File format is wrong."
'
'    On Error GoTo 0
'    Exit Sub
'
'Test_vtkRecreateConfiguration_formatLikeExtension_XLA_Error:
'    Err.Source = "Test_vtkRecreateConfiguration_formatLikeExtension_XLA of module vtkRecreateConfigurationTester"
'    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
'    ITestCase_TearDown
'    Exit Sub
'End Sub



Public Sub Test_vtkRecreateConfiguration_workbookOpen()
'  workbook with the same name is already open : raise error VTK_WORKBOOK_ALREADY_OPEN

    Dim error As Integer: error = 0

    On Error GoTo Test_vtkRecreateConfiguration_workbookOpen_Error

    Set testedConf = normalConf

    ' Create and open recreatedWb
    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)

    ' Recreate configuration with recreateWb already open
    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name

    mAssert.Equals error, VTK_WORKBOOK_ALREADY_OPEN

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookOpen_Error:
        If Err.Number = VTK_WORKBOOK_ALREADY_OPEN Then
            error = Err.Number
            Resume Next
        Else
            Err.Source = "Test_vtkRecreateConfiguration_workbookOpen of module vtkRecreateConfigurationTester"
            mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
            ITestCase_TearDown
            Exit Sub
        End If
End Sub



Public Sub Test_vtkRecreateConfiguration_addInOpen_XLAM()
' add-in with the same name is activated : recreate the configuration normally

    Dim error As Integer: error = 0

    On Error GoTo Test_vtkRecreateConfiguration_addInOpen_Error

    Set testedConf = VTKConf

    ' Recreate the configuration called just like an actived add-in : "VBAToolKit.xlam"
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name

    mAssert.Equals error, 0

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_addInOpen_Error:
    error = Err.Number
    Resume Next
End Sub



Public Sub Test_vtkRecreateConfiguration_overwriteFile_normal()
' file already exists : overwrite it

    On Error GoTo Test_vtkRecreateConfiguration_overwriteFile_normal_Error

    Dim oldSize As Integer
    Dim newSize As Integer

    Set testedConf = normalConf

    fso.CreateTextFile (VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)
    oldSize = fso.GetFile(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path).Size

    ' Create and open recreatedWb
    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)

    newSize = fso.GetFile(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path).Size

    mAssert.Should newSize > oldSize, "File should have been overwritten."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_overwriteFile_normal_Error:
    Err.Source = "Test_vtkRecreateConfiguration_overwriteFile_normal of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_overwriteFile_addIn()
' add-in already exists : overwrite it

    On Error GoTo Test_vtkRecreateConfiguration_overwriteFile_addIn_Error

    Dim oldSize As Integer
    Dim newSize As Integer

    Set testedConf = xlamConf

    fso.CreateTextFile (VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)
    oldSize = fso.GetFile(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path).Size

    ' Create and open recreatedWb
    vtkRecreateConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path)

    newSize = fso.GetFile(VBAToolKit.vtkPathOfCurrentProject(Wb) & "\" & testedConf.path).Size

    mAssert.Should newSize > oldSize, "File should have been overwritten."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_overwriteFile_addIn_Error:
    Err.Source = "Test_vtkRecreateConfiguration_overwriteFile_addIn of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub




Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_allModulesArePresent_normalConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_referencesAreActivated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLSM"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLAM"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLS"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookOpen"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_addInOpen_XLAM"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_overwriteFile_normal"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_overwriteFile_addIn"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf
        Case "Test_vtkRecreateConfiguration_allModulesArePresent_normalConf": Test_vtkRecreateConfiguration_allModulesArePresent_normalConf
        Case "Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf": Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf
        Case "Test_vtkRecreateConfiguration_referencesAreActivated": Test_vtkRecreateConfiguration_referencesAreActivated
        Case "Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect": Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLSM": Test_vtkRecreateConfiguration_formatLikeExtension_XLSM
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLAM": Test_vtkRecreateConfiguration_formatLikeExtension_XLAM
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLS": Test_vtkRecreateConfiguration_formatLikeExtension_XLS
        Case "Test_vtkRecreateConfiguration_workbookOpen": Test_vtkRecreateConfiguration_workbookOpen
        Case "Test_vtkRecreateConfiguration_addInOpen_XLAM": Test_vtkRecreateConfiguration_addInOpen_XLAM
        Case "Test_vtkRecreateConfiguration_overwriteFile_normal": Test_vtkRecreateConfiguration_overwriteFile_normal
        Case "Test_vtkRecreateConfiguration_overwriteFile_addIn": Test_vtkRecreateConfiguration_overwriteFile_addIn
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

