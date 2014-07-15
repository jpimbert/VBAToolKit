VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkRecreateConfExcelTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkRecreateConfExcelTester
' Author    : Lucas Vitorino
' Purpose   : Test the vtkRecreateConfiguration function with Excel Conf Manager
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

Private fso As FileSystemObject

Private Wb As Workbook
Private recreatedWb As Workbook
Private testProject As vtkProject
Private normalConf As vtkConfiguration
Private emptyConf As vtkConfiguration
Private notExistingPathConf1 As vtkConfiguration
Private notExistingPathConf2 As vtkConfiguration
Private xlsConf As vtkConfiguration
Private xlamConf As vtkConfiguration
Private xlaConf As vtkConfiguration
Private multitypeConf As vtkConfiguration
Private noTemplateConf As vtkConfiguration
Private VTKConf As vtkConfiguration
Private VTKProjectConf As vtkConfiguration
Private testedConf As vtkConfiguration

Private cm As vtkConfigurationManager

Private Const testProjectName = "TestProjectForRecreateConf"
Private Const normalConfName = "TestProject_normalConf"
Private Const normalConfTemplate = "TestProjectTemplate.xlsx"
Private Const emptyConfName = "TestProject_emptyConf"
Private Const notExistingPathConf1Name = "TestProject_notExistingPathConf1"
Private Const notExistingPathConf2Name = "TestProject_notExistingPathConf2"
Private Const xlsConfName = "TestProject_xlsConf"
Private Const xlamConfName = "TestProject_xlamConf"
Private Const xlaConfName = "TestProject_xlaConf"
Private Const multitypeConfName = "TestProject_multitypeConf"
Private Const noTemplateConfName = "TestProject_NoTemplate"
Private Const VTKConfName = "TestProject_VTKConf"
Private Const VTKProjectConfName = "VBAToolKit"

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
    
    Set fso = New FileSystemObject

    Set testProject = vtkProjectForName(testProjectName)
    
    Dim testProjectFolder As String
    
    testProjectFolder = fso.BuildPath(VBAToolKit.vtkTestPath, testProject.projectName)
    vtkCreateTreeFolder testProjectFolder
    
    Set Wb = getTestFileFromTemplate(testProject.workbookDEVName, _
             destinationName:=fso.BuildPath(testProject.projectName, testProject.projectDEVStandardRelativePath), _
             openExcel:=True)
    Set cm = vtkConfigurationManagerForProject(testProject.projectName)
    
    getTestFileFromTemplate fileName:=normalConfTemplate, _
         destinationName:=fso.BuildPath(testProject.projectName, "Templates" & "\" & normalConfTemplate), _
         openExcel:=False

    Set normalConf = cm.configurations(normalConfName)
    Set emptyConf = cm.configurations(emptyConfName)
    Set notExistingPathConf1 = cm.configurations(notExistingPathConf1Name)
    Set notExistingPathConf2 = cm.configurations(notExistingPathConf2Name)
    Set xlsConf = cm.configurations(xlsConfName)
    Set xlamConf = cm.configurations(xlamConfName)
    Set xlaConf = cm.configurations(xlaConfName)
    Set multitypeConf = cm.configurations(multitypeConfName)
    Set noTemplateConf = cm.configurations(noTemplateConfName)
    Set VTKConf = cm.configurations(VTKConfName)
    Set VTKProjectConf = cm.configurations(VTKProjectConfName)

    On Error GoTo 0
    Exit Sub

err_handler:
    Debug.Print "SETUP : Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
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
'   - modules have not been exported
'       - normal : raise error VTK_NO_SOURCE_FILES
'       - empty conf : raise no error
'   - workbook is to be created in a path that does not exist yet
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
'   - VBProject with the same name is already open : raise error VTK_VBPROJECT_ALREADY_OPEN
'   - add-in with the same name is activated : export the file normally
'   - workbook with the same name already exists in the same path : overwrite the file
'       - normal conf
'       - xlam

' --------------------------------------------------------
'               PRIVATE UTILITY FUNCTIONS
' --------------------------------------------------------

Private Sub exportModulesAndRecreateWb(Wb As Workbook, testProject As vtkProject, testedConf As vtkConfiguration)
    ' It is mandatory to create the source files before recreating a configuration
    On Error GoTo exportModulesAndRecreateWb_Error

    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path)

    On Error GoTo 0
    Exit Sub

exportModulesAndRecreateWb_Error:
    Err.Source = "exportModulesAndRecreateWb of module vtkRecreateConfigurationTester"
    Err.Raise Err.Number, Err.source, Err.Description
    Exit Sub
End Sub

Private Sub verifyV11Properties(Wb As Workbook, confNumber As Integer)
    mAssert.Equals Wb.VBProject.name, "ProjectName" & confNumber, "Expected project name"
    mAssert.Equals Wb.BuiltinDocumentProperties("Title").Value, "ProjectName" & confNumber, "Expected Workbook title"
    mAssert.Equals Wb.BuiltinDocumentProperties("Comments").Value, "Comment for Workbook #" & confNumber, "Expected Workbook comment"
End Sub

' --------------------------------------------------------
'                       TESTS
' --------------------------------------------------------

Public Sub Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf()
' workbook is created in the good path

    On Error GoTo Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf_Error

    Set testedConf = normalConf

    exportModulesAndRecreateWb wb, testProject, testedConf

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The configuration should have been created in the designated path."
    verifyV11Properties recreatedWb, testedConf.Number
    
    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf_Error:
    Err.Source = "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfigurations_normalConf()
' Same test than previous one but call of main function instead of utility function
' The normal conf contains a template ; check the use of this template

    On Error GoTo Test_vtkRecreateConfiguration_normalConf_Error

    Set testedConf = normalConf

    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    Dim confNames As New Collection
    confNames.Add testedConf.name
    vtkRecreateConfigurations cm, confNames
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path)

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The configuration should have been created in the designated path."
    
    mAssert.Equals recreatedWb.Sheets("TestProject").Range("A1").Value, "TestProjectTemplate", "The configuration should have use a template"

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_normalConf_Error:
    Err.Source = "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf()
' workbook is created in the good path

    On Error GoTo Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf_Error

    Set testedConf = emptyConf

    exportModulesAndRecreateWb wb, testProject, testedConf

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The configuration should have been created in the designated path."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf_Error:
    Err.Source = "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf()
' workbook is created in the good path

    On Error GoTo Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf_Error

    Set testedConf = multitypeConf

    exportModulesAndRecreateWb wb, testProject, testedConf

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The configuration should have been created in the designated path."
    verifyV11Properties recreatedWb, testedConf.Number

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf_Error:
    Err.Source = "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_modulesNotYetExported_normalConf()
' modules have not been exported, normal conf : raise error VTK_NO_SOURCE_FILES

    On Error GoTo Test_vtkRecreateConfiguration_modulesNotYetExported_normalConf_Error

    Dim error As Integer: error = 0

    Set testedConf = normalConf

    vtkRecreateConfiguration testProject.projectName, testedConf.name

    mAssert.Equals error, VTK_NO_SOURCE_FILES

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_modulesNotYetExported_normalConf_Error:
    error = Err.Number
    Resume Next
End Sub




Public Sub Test_vtkRecreateConfiguration_modulesNotYetExported_emptyConf()
' modules have not been exported, empty conf : don't raise error, create the file normally

    On Error GoTo Test_vtkRecreateConfiguration_modulesNotYetExported_emptyConf_Error

    Dim error As Integer: error = 0

    Set testedConf = emptyConf

    vtkRecreateConfiguration testProject.projectName, testedConf.name

    mAssert.Equals error, 0, _
                   "Error " & Err.Number & " (" & Err.Description & ") occured."

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The file shoud have been created normally as configuration is empty."
    verifyV11Properties recreatedWb, testedConf.Number

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_modulesNotYetExported_emptyConf_Error:
    error = Err.Number
    Resume Next
End Sub


Public Sub Test_vtkRecreateConfiguration_pathDoesNotExistYet1()
' normal conf, 1-level path does not exist yet : don't raise error, create the file normally

    On Error GoTo Test_vtkRecreateConfiguration_pathDoesNotExistYet1_Error

    Dim error As Integer: error = 0

    Set testedConf = notExistingPathConf1

    exportModulesAndRecreateWb wb, testProject, testedConf

    mAssert.Equals error, 0, _
                    "Error " & error & " (" & Err.Description & ") occured. It shouldn't have."

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The file shoud have been created normally."
    verifyV11Properties recreatedWb, testedConf.Number

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_pathDoesNotExistYet1_Error:
    error = Err.Number
    Resume Next
End Sub


Public Sub Test_vtkRecreateConfiguration_pathDoesNotExistYet2()
' normal conf, 2-level path does not exist yet : raise error VTK_WRONG_FILE_PATH, don't create the file

    On Error GoTo Test_vtkRecreateConfiguration_pathDoesNotExistYet2_Error

    Dim error As Integer: error = 0

    Set testedConf = notExistingPathConf2

    exportModulesAndRecreateWb wb, testProject, testedConf

    mAssert.Equals error, 0, _
                    "Error " & error & " (" & Err.Description & ") occured. It shouldn't have."

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The file shoud have been created normally."
    verifyV11Properties recreatedWb, testedConf.Number
    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_pathDoesNotExistYet2_Error:
    error = Err.Number
    Resume Next
End Sub


Public Sub Test_vtkRecreateConfiguration_allModulesArePresent_normalConf()
' all modules of the configuration are present

    Dim error As Integer: error = 0

    On Error GoTo Test_vtkRecreateConfiguration_allModulesArePresent_normalConf_Error

    Set testedConf = normalConf

    exportModulesAndRecreateWb wb, testProject, testedConf

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
            Exit Sub
        End If
End Sub



Public Sub Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf()
' all modules of the configuration are present

    Dim error As Integer: error = 0

    On Error GoTo Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf_Error

    Set testedConf = multitypeConf

    exportModulesAndRecreateWb wb, testProject, testedConf

    Dim mo As vtkModule
    Dim vbCmp As VBComponent
    For Each mo In testedConf.modules
        ' Just a dummy access attempt : will throw error 9 (subscript out of range) if the module doesn't exist
        Set vbCmp = recreatedWb.VBProject.VBComponents(mo.name)
    Next

    ' If there is no error, an assert instruction has to be executed
    mAssert.Should True

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf_Error:
        If Err.Number = 9 Then
            error = Err.Number
            mAssert.Should False, "Module " & mo.name & "is missing."
            Resume Next
        Else
            Err.Source = "Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf of module vtkRecreateConfigurationTester"
            mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
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

    exportModulesAndRecreateWb wb, testProject, testedConf

    ' Listing expected references in a collection
    numRef = 4
    ReDim guidArray(1 To numRef) As String
    guidArray(1) = "{420B2830-E718-11CF-893D-00A0C9054228}" 'Scripting : Microsoft scripting runtime
    guidArray(2) = "{0002E157-0000-0000-C000-000000000046}" ' VBIDE : Microsoft visual basic for applications extensibility 5.3
    guidArray(3) = "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}" ' Shell32 : Microsoft Shell Controls and Automation+
    guidArray(4) = "{F5078F18-C551-11D3-89B9-0000F81FE221}" ' MSXML2 : Microsoft XML V5.0

    ' For each GUID in the array, check if it is in the project references
    For i = 1 To numRef
        refIsActive = False
        For Each ref In recreatedWb.VBProject.references
            If ref.GUID = guidArray(i) Then refIsActive = True
        Next

        mAssert.Should refIsActive, "The reference with GUID = " & guidArray(i) & " is not activated : it should."
    Next

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_referencesAreActivated_Error:
    Err.Source = "Test_vtkRecreateConfiguration_referencesAreActivated of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect()
' name of the vbproject is correct : same name as the recreated configuration

    On Error GoTo Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect_Error

    Set testedConf = normalConf

    ' Set recreatedWb and open it
    exportModulesAndRecreateWb wb, testProject, testedConf

    mAssert.Equals recreatedWb.VBProject.name, "ProjectName" & testedConf.Number, _
                   "The created project should have the same name as the recreated configuration."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect_Error:
    Err.Source = "Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_formatLikeExtension_XLSM()
' test the xlsm format

    On Error GoTo Test_vtkRecreateConfiguration_formatLikeExtension_XLSM_Error

    Set testedConf = normalConf

    exportModulesAndRecreateWb wb, testProject, testedConf

    mAssert.Equals recreatedWb.FileFormat, _
                   xlOpenXMLWorkbookMacroEnabled, _
                   "File format is wrong."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_formatLikeExtension_XLSM_Error:
    Err.Source = "Test_vtkRecreateConfiguration_formatLikeExtension_XLSM of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_formatLikeExtension_XLAM()
' test the xlam format

    On Error GoTo Test_vtkRecreateConfiguration_formatLikeExtension_XLAM_Error

    Set testedConf = xlamConf

    exportModulesAndRecreateWb wb, testProject, testedConf

    mAssert.Equals recreatedWb.FileFormat, _
                   xlOpenXMLAddIn, _
                   "File format is wrong."
    verifyV11Properties recreatedWb, testedConf.Number

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_formatLikeExtension_XLAM_Error:
    Err.Source = "Test_vtkRecreateConfiguration_formatLikeExtension_XLAM of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_formatLikeExtension_XLS()
' test the xls format

    On Error GoTo Test_vtkRecreateConfiguration_formatLikeExtension_XLS_Error

    Set testedConf = xlsConf

    exportModulesAndRecreateWb wb, testProject, testedConf

    mAssert.Equals recreatedWb.FileFormat, _
                   xlExcel8, _
                   "File format is wrong."
    verifyV11Properties recreatedWb, testedConf.Number

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_formatLikeExtension_XLS_Error:
    Err.Source = "Test_vtkRecreateConfiguration_formatLikeExtension_XLS of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfiguration_formatLikeExtension_XLA()
' test the xla format

    On Error GoTo Test_vtkRecreateConfiguration_formatLikeExtension_XLA_Error

    Set testedConf = xlaConf

    exportModulesAndRecreateWb wb, testProject, testedConf

    mAssert.Equals recreatedWb.FileFormat, _
                   xlAddIn, _
                   "File format is wrong."
    verifyV11Properties recreatedWb, testedConf.Number

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_formatLikeExtension_XLA_Error:
    Err.Source = "Test_vtkRecreateConfiguration_formatLikeExtension_XLA of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub

Public Sub Test_vtkRecreateConfiguration_workbookOpen()
'  workbook with the same name is already open : raise error VTK_WORKBOOK_ALREADY_OPEN

    Dim error As Integer: error = 0

    On Error GoTo Test_vtkRecreateConfiguration_workbookOpen_Error

    Set testedConf = normalConf

    ' Create and open recreatedWb
    exportModulesAndRecreateWb wb, testProject, testedConf

    ' Recreate configuration with recreateWb already open
    vtkRecreateConfiguration testProject.projectName, testedConf.name

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
            Exit Sub
        End If
End Sub



Public Sub Test_vtkRecreateConfiguration_addInOpen_XLAM()
' add-in with the same name is activated : recreate the configuration normally

    Dim error As Integer: error = 0

    On Error GoTo Test_vtkRecreateConfiguration_addInOpen_Error

    Set testedConf = VTKConf

    ' Recreate the configuration called just like an actived add-in : "VBAToolKit.xlam"
    ' NB : the recreated workbook is not opened
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration testProject.projectName, testedConf.name

    mAssert.Equals error, 0

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_addInOpen_Error:
    error = Err.Number
    Resume Next
End Sub



Public Sub Test_vtkRecreateConfiguration_workbookAndVbProjectAlreadyOpen()
' workbook and vbproject with the same name already open -> there shouldn't be errors

    Dim error As Integer: error = 0

    On Error GoTo Test_vtkRecreateConfiguration_workbookAndVbProjectAlreadyOpen_Error

    Set testedConf = VTKProjectConf

    ' Recreate the configuration called just like an actived add-in : "VBAToolKit.xlam"
    ' NB : the recreated workbook is not opened
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    vtkRecreateConfiguration testProject.projectName, testedConf.name

    mAssert.Equals error, 0

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookAndVbProjectAlreadyOpen_Error:
    error = Err.Number
    Resume Next
End Sub

Public Sub Test_vtkRecreateConfiguration_unreachableTemplate()
' workbook with an unexisting template

    Dim error As Integer: error = 0

    On Error GoTo M_Error

    Set testedConf = noTemplateConf
    vtkExportConfiguration Wb.VBProject, testProject.projectName, testedConf.name
    
   On Error Resume Next
    vtkRecreateConfiguration testProject.projectName, testedConf.name
    mAssert.Equals Err.Number, VTK_TEMPLATE_NOT_FOUND, "Error returned by vtkRecreateConfiguration"
   On Error GoTo 0
    Exit Sub

M_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ")"
End Sub


Public Sub Test_vtkRecreateConfiguration_overwriteFile_normal()
' file already exists : overwrite it

    On Error GoTo Test_vtkRecreateConfiguration_overwriteFile_normal_Error

    Dim oldSize As Integer
    Dim newSize As Integer

    Set testedConf = normalConf

    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path)
    oldSize = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path).Size

    ' Create and open recreatedWb
    exportModulesAndRecreateWb wb, testProject, testedConf

    newSize = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path).Size

    mAssert.Should newSize > oldSize, "File should have been overwritten."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_overwriteFile_normal_Error:
    Err.Source = "Test_vtkRecreateConfiguration_overwriteFile_normal of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_overwriteFile_addIn()
' add-in already exists : overwrite it

    On Error GoTo Test_vtkRecreateConfiguration_overwriteFile_addIn_Error

    Dim oldSize As Integer
    Dim newSize As Integer

    Set testedConf = xlamConf

    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path)
    oldSize = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path).Size

    ' Create and open recreatedWb
    exportModulesAndRecreateWb wb, testProject, testedConf

    newSize = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path).Size

    mAssert.Should newSize > oldSize, "File should have been overwritten."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_overwriteFile_addIn_Error:
    Err.Source = "Test_vtkRecreateConfiguration_overwriteFile_addIn of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfigurations_normalConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_modulesNotYetExported_normalConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_modulesNotYetExported_emptyConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_pathDoesNotExistYet1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_pathDoesNotExistYet2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_allModulesArePresent_normalConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_referencesAreActivated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLSM"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLAM"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLS"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLA"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookOpen"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_addInOpen_XLAM"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookAndVbProjectAlreadyOpen"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_unreachableTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_overwriteFile_normal"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_overwriteFile_addIn"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf
        Case "Test_vtkRecreateConfigurations_normalConf": Test_vtkRecreateConfigurations_normalConf
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf
        Case "Test_vtkRecreateConfiguration_modulesNotYetExported_normalConf": Test_vtkRecreateConfiguration_modulesNotYetExported_normalConf
        Case "Test_vtkRecreateConfiguration_modulesNotYetExported_emptyConf": Test_vtkRecreateConfiguration_modulesNotYetExported_emptyConf
        Case "Test_vtkRecreateConfiguration_pathDoesNotExistYet1": Test_vtkRecreateConfiguration_pathDoesNotExistYet1
        Case "Test_vtkRecreateConfiguration_pathDoesNotExistYet2": Test_vtkRecreateConfiguration_pathDoesNotExistYet2
        Case "Test_vtkRecreateConfiguration_allModulesArePresent_normalConf": Test_vtkRecreateConfiguration_allModulesArePresent_normalConf
        Case "Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf": Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf
        Case "Test_vtkRecreateConfiguration_referencesAreActivated": Test_vtkRecreateConfiguration_referencesAreActivated
        Case "Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect": Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLSM": Test_vtkRecreateConfiguration_formatLikeExtension_XLSM
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLAM": Test_vtkRecreateConfiguration_formatLikeExtension_XLAM
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLS": Test_vtkRecreateConfiguration_formatLikeExtension_XLS
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLA": Test_vtkRecreateConfiguration_formatLikeExtension_XLA
        Case "Test_vtkRecreateConfiguration_workbookOpen": Test_vtkRecreateConfiguration_workbookOpen
        Case "Test_vtkRecreateConfiguration_addInOpen_XLAM": Test_vtkRecreateConfiguration_addInOpen_XLAM
        Case "Test_vtkRecreateConfiguration_workbookAndVbProjectAlreadyOpen": Test_vtkRecreateConfiguration_workbookAndVbProjectAlreadyOpen
        Case "Test_vtkRecreateConfiguration_unreachableTemplate": Test_vtkRecreateConfiguration_unreachableTemplate
        Case "Test_vtkRecreateConfiguration_overwriteFile_normal": Test_vtkRecreateConfiguration_overwriteFile_normal
        Case "Test_vtkRecreateConfiguration_overwriteFile_addIn": Test_vtkRecreateConfiguration_overwriteFile_addIn
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

