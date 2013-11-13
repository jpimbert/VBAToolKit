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
Private testProject As vtkProject
Private testProjectFolder As String
Private normalConf As vtkConfiguration
Private emptyConf As vtkConfiguration
Private notExistingPathConf1 As vtkConfiguration
Private notExistingPathConf2 As vtkConfiguration
Private xlsConf As vtkConfiguration
Private xlamConf As vtkConfiguration
Private xlaConf As vtkConfiguration
Private multitypeConf As vtkConfiguration
Private vtkConf As vtkConfiguration
Private VTKProjectConf As vtkConfiguration
Private testedConf As vtkConfiguration
Private faultyRefConf As vtkConfiguration
Private xlsmTemplateConf As vtkConfiguration
Private xlamTemplateConf As vtkConfiguration
Private xlaTemplateConf As vtkConfiguration
Private xlsTemplateConf As vtkConfiguration

Private cm As vtkConfigurationManager

Private Const testProjectName = "TestProjectForRecreateConfiguration"
Private Const normalConfName = "TestProject_normalConf"
Private Const emptyConfName = "TestProject_emptyConf"
Private Const notExistingPathConf1Name = "TestProject_notExistingPathConf1"
Private Const notExistingPathConf2Name = "TestProject_notExistingPathConf2"
Private Const xlsConfName = "TestProject_xlsConf"
Private Const xlamConfName = "TestProject_xlamConf"
Private Const xlaConfName = "TestProject_xlaConf"
Private Const multitypeConfName = "TestProject_multitypeConf"
Private Const VTKConfName = "TestProject_VTKConf"
Private Const faultyRefConfName = "TestProject_faultyRefConf"
Private Const xlsmTemplateConfName = "TestProject_xlsmTemplateConf"
Private Const xlamTemplateConfName = "TestProject_xlamTemplateConf"
Private Const xlaTemplateConfName = "TestProject_xlaTemplateConf"
Private Const xlsTemplateConfName = "TestProject_xlsTemplateConf"
Private Const VTKProjectConfName = "VBAToolKit"

Private Const rememberedProjectsSheetName As String = "VBAToolKitTestProjects.xml"

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
    
    ' TEMPORARY ---------------
    VBAToolKit.resetTestFolder
    vtkResetProjects
    vtkResetConfigurationManagers
    ' -------------------------
    
    ' Initialize useful variables and environment
    Set testProject = vtkProjectForName(testProjectName)
    testProjectFolder = fso.BuildPath(VBAToolKit.vtkTestPath, testProject.projectName)
    vtkCreateTreeFolder testProjectFolder

    ' Create the files we'll need for remembering projects
    createRememberedProjectsXMLSheet sheetPath:=fso.BuildPath(VBAToolKit.vtkTestPath, rememberedProjectsSheetName)
    
    ' Copy the sheet of the new project in the relevant folder
    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, testProject.XMLConfigurationName), _
             destination:=fso.BuildPath(testProjectFolder, testProject.XMLConfigurationStandardRelativePath)
        
    ' Add the TestProject to the list of remembered projects
    xmlRememberedProjectsFullPath = fso.BuildPath(VBAToolKit.vtkTestPath, rememberedProjectsSheetName)
    
    ' Add "TestProject" to the list of remembered projects
    vtkAddRememberedProject projectName:=testProject.projectName, _
                            rootFolder:=fso.BuildPath(VBAToolKit.vtkTestPath, testProject.projectName), _
                            xmlRelativePath:=testProject.XMLConfigurationStandardRelativePath
                                            
    ' Set all the configuration objects
    Set cm = vtkConfigurationManagerForProject(testProject.projectName)
    Set normalConf = cm.configurations(normalConfName)
    Set emptyConf = cm.configurations(emptyConfName)
    Set notExistingPathConf1 = cm.configurations(notExistingPathConf1Name)
    Set notExistingPathConf2 = cm.configurations(notExistingPathConf2Name)
    Set xlsConf = cm.configurations(xlsConfName)
    Set xlamConf = cm.configurations(xlamConfName)
    Set xlaConf = cm.configurations(xlaConfName)
    Set multitypeConf = cm.configurations(multitypeConfName)
    Set vtkConf = cm.configurations(VTKConfName)
    Set VTKProjectConf = cm.configurations(VTKProjectConfName)
    Set faultyRefConf = cm.configurations(faultyRefConfName)
    Set xlsmTemplateConf = cm.configurations(xlsmTemplateConfName)
    Set xlamTemplateConf = cm.configurations(xlamTemplateConfName)
    Set xlaTemplateConf = cm.configurations(xlaTemplateConfName)
    Set xlsTemplateConf = cm.configurations(xlsTemplateConfName)
    
    ' Copy the workbook templates
    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, "xlsmTemplateWorkbook.xlsm"), _
             destination:=fso.BuildPath(testProjectFolder, "xlsmTemplateWorkbook.xlsm")
             
    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, "xlamTemplateWorkbook.xlam"), _
             destination:=fso.BuildPath(testProjectFolder, "xlamTemplateWorkbook.xlam")

    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, "xlaTemplateWorkbook.xla"), _
             destination:=fso.BuildPath(testProjectFolder, "xlaTemplateWorkbook.xla")
             
    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, "xlsTemplateWorkbook.xls"), _
             destination:=fso.BuildPath(testProjectFolder, "xlsTemplateWorkbook.xls")


    On Error GoTo 0
    Exit Sub

err_handler:
    Debug.Print "SETUP : Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub

Private Sub ITestCase_TearDown()
    On Error Resume Next
    vtkCloseAndKillWorkbook Wb:=recreatedWb
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

Private Sub getModules()
    ' Copy the modules (mandatory to recreate the configuration)
    On Error GoTo getModules_Error

    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, "RecreateConfigurationModules\TestModule1.bas"), _
             destination:=fso.BuildPath(testProjectFolder, "TestModule1.bas")
             
    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, "RecreateConfigurationModules\TestModule2.bas"), _
             destination:=fso.BuildPath(testProjectFolder, "TestModule2.bas")
    
    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, "RecreateConfigurationModules\TestClassModule1.cls"), _
             destination:=fso.BuildPath(testProjectFolder, "TestClassModule1.cls")
    
    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, "RecreateConfigurationModules\TestFormModule1.frm"), _
             destination:=fso.BuildPath(testProjectFolder, "TestFormModule1.frm")
             
    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, "RecreateConfigurationModules\TestFormModule1.frx"), _
             destination:=fso.BuildPath(testProjectFolder, "TestFormModule1.frx")

    On Error GoTo 0
    Exit Sub

getModules_Error:
    Err.Source = "vtkRecreateConfigurationTester::getModules"
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub

End Sub


Private Sub getModulesAndRecreateWb(testProject As vtkProject, testedConf As vtkConfiguration)
    ' It is mandatory to create the source files before recreating a configuration
    On Error GoTo getModulesAndRecreateWb_Error
    
    ' Copy the modules (mandatory to recreate the configuration)
    getModules

    ' Recreate the configuration
    vtkRecreateConfiguration testProject.projectName, testedConf.name
    Set recreatedWb = Workbooks.Open(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path)

    On Error GoTo 0
    Exit Sub

getModulesAndRecreateWb_Error:
    Err.Source = "getModulesAndRecreateWb  of module vtkRecreateConfigurationTester"
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub


' --------------------------------------------------------
'                       TESTS
' --------------------------------------------------------
'
Public Sub Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf()
' workbook is created in the good path

    On Error GoTo Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf_Error

    Set testedConf = normalConf

    getModulesAndRecreateWb testProject, testedConf

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The configuration should have been created in the designated path."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf_Error:
    Err.Source = "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf()
' workbook is created in the good path

    On Error GoTo Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf_Error

    Set testedConf = emptyConf

    getModulesAndRecreateWb testProject, testedConf

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

    getModulesAndRecreateWb testProject, testedConf

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The configuration should have been created in the designated path."

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

    getModulesAndRecreateWb testProject, testedConf

    mAssert.Equals error, 0, _
                    "Error " & error & " (" & Err.Description & ") occured. It shouldn't have."

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The file shoud have been created normally."

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

    getModulesAndRecreateWb testProject, testedConf

    mAssert.Equals error, 0, _
                    "Error " & error & " (" & Err.Description & ") occured. It shouldn't have."

    mAssert.Should fso.FileExists(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path), _
                   "The file shoud have been created normally."
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

    getModulesAndRecreateWb testProject, testedConf

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

    getModulesAndRecreateWb testProject, testedConf

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
            mAssert.Should False, "Module " & mo.name & " is missing."
            Resume Next
        Else
            Err.Source = "Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf of module vtkRecreateConfigurationTester"
            mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
            Exit Sub
        End If
End Sub


Public Sub Test_vtkRecreateConfiguration_referencesAreActivated()
' references are activated

    On Error GoTo Test_vtkRecreateConfiguration_referencesAreActivated_Error

    Set testedConf = normalConf

    Dim numRef As Integer
    Dim ref As Reference
    Dim i As Integer
    Dim refIsActive As Boolean
    Dim guidArray() As String

    getModulesAndRecreateWb testProject, testedConf

    ' Listing expected references in a collection
    numRef = 2
    ReDim guidArray(1 To numRef) As String
    guidArray(1) = "{420B2830-E718-11CF-893D-00A0C9054228}" 'Scripting : Microsoft scripting runtime
    guidArray(2) = "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}" ' Shell32 : Microsoft Shell Controls and Automation+

    ' For each GUID in the array, check if it is in the project references
    For i = 1 To numRef
        refIsActive = False
        For Each ref In recreatedWb.VBProject.references
            If ref.guid = guidArray(i) Then refIsActive = True
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


Public Sub Test_vtkRecreateConfiguration_badReference()
' appropriate error is thrown

    On Error GoTo Test_vtkRecreateConfiguration_badReference_Error

    Set testedConf = faultyRefConf

    ' A reference with a faulty guid has been set
    getModulesAndRecreateWb testProject, testedConf

    mAssert.Should False, "An error should have been raised."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_badReference_Error:
    Err.Source = "Test_vtkRecreateConfiguration_badReference of module vtkRecreateConfigurationTester"
    mAssert.Equals Err.Number, VTK_REFERENCE_ERROR
    Exit Sub
End Sub


Public Sub Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect()
' name of the vbproject is correct : same name as the recreated configuration

    On Error GoTo Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect_Error

    Set testedConf = normalConf

    ' Set recreatedWb and open it
    getModulesAndRecreateWb testProject, testedConf

    mAssert.Equals recreatedWb.VBProject.name, testedConf.name, _
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

    getModulesAndRecreateWb testProject, testedConf

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

    getModulesAndRecreateWb testProject, testedConf

    mAssert.Equals recreatedWb.FileFormat, _
                   xlOpenXMLAddIn, _
                   "File format is wrong."

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

    getModulesAndRecreateWb testProject, testedConf

    mAssert.Equals recreatedWb.FileFormat, _
                   xlExcel8, _
                   "File format is wrong."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_formatLikeExtension_XLS_Error:
    Err.Source = "Test_vtkRecreateConfiguration_formatLikeExtension_XLS of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkRecreateConfiguration_formatLikeExtension_XLA()
' test the xla format
' test failing, don 't know why

    On Error GoTo Test_vtkRecreateConfiguration_formatLikeExtension_XLA_Error

    Set testedConf = xlaConf

    getModulesAndRecreateWb testProject, testedConf

    mAssert.Equals recreatedWb.FileFormat, _
                   xlAddIn, _
                   "File format is wrong."

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
    getModulesAndRecreateWb testProject, testedConf

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

    Set testedConf = vtkConf

    ' Recreate the configuration called just like an actived add-in : "VBAToolKit.xlam"
    ' NB : the recreated workbook is not opened
    getModules
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
    getModules
    vtkRecreateConfiguration testProject.projectName, testedConf.name

    mAssert.Equals error, 0

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_workbookAndVbProjectAlreadyOpen_Error:
    error = Err.Number
    Resume Next
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
    getModulesAndRecreateWb testProject, testedConf

    newSize = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path).Size

    mAssert.Should newSize > oldSize, "File should have been overwritten."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_overwriteFile_normal_Error:
    Err.Source = "Test_vtkRecreateConfiguration_overwriteFile_normal of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub


'
Public Sub Test_vtkRecreateConfiguration_overwriteFile_addIn()
' add-in already exists : overwrite it

    On Error GoTo Test_vtkRecreateConfiguration_overwriteFile_addIn_Error

    Dim oldSize As Integer
    Dim newSize As Integer

    Set testedConf = xlamConf

    fso.CreateTextFile (VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path)
    oldSize = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path).Size

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    newSize = fso.GetFile(VBAToolKit.vtkTestPath & "\" & testProject.projectName & "\" & testedConf.path).Size

    mAssert.Should newSize > oldSize, "File should have been overwritten."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_overwriteFile_addIn_Error:
    Err.Source = "Test_vtkRecreateConfiguration_overwriteFile_addIn of module vtkRecreateConfigurationTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub


Public Sub Test_vtkRecreateConfiguration_title()
' the title of the workbook should be right

    On Error GoTo Test_vtkRecreateConfiguration_title_Error

    Set testedConf = normalConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    mAssert.Equals recreatedWb.BuiltinDocumentProperties("Title"), "Normal Conf"

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_title_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_title"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkRecreateConfiguration_comment()
' the comment of the workbook should be right

    On Error GoTo Test_vtkRecreateConfiguration_comment_Error

    Set testedConf = normalConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    mAssert.Equals recreatedWb.BuiltinDocumentProperties("Comments"), "This conf is normal"

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_comment_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_comment"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub Test_vtkRecreateConfiguration_templateXlsmVBProject()
' ensure the configuration is well created with an xlsm template, and that the VBProjec name is set
' according to the template

    On Error GoTo Test_vtkRecreateConfiguration_templateXlsmVBProject_Error

    Set testedConf = xlsmTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the VBProjectName
    mAssert.Equals recreatedWb.VBProject.name, "xlsmTemplateProject"

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlsmVBProject_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlsmVBProject"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkRecreateConfiguration_templateXlsmModule()
' ensure the configuration is well created with an xlsm template, and that the module of the template
' exists in the recreated Workbook

    On Error GoTo Test_vtkRecreateConfiguration_templateXlsmmodule_Error

    Set testedConf = xlsmTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the presence of the module from the template
    mAssert.Should Not recreatedWb.VBProject.VBComponents("xlsmTemplateModule") Is Nothing

    mAssert.Should True, "The above shouldn't have thrown any error."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlsmmodule_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlsmmodule"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub Test_vtkRecreateConfiguration_templateXlsmWorksheet()
' ensure the configuration is well created with an xlsm template, and that the worksheet present in the template
' are present too in the recreatedWb

    On Error GoTo Test_vtkRecreateConfiguration_templateXlsmWorksheet_Error

    Set testedConf = xlsmTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the presence of the worksheet from the template
    mAssert.Should Not recreatedWb.Worksheets("xlsmTemplateSheet") Is Nothing

    mAssert.Should True, "The above shouldn't have thrown any error."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlsmWorksheet_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlsmWorksheet"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkRecreateConfiguration_templateXlamVBProject()
' ensure the configuration is well created with an xlam template, and that the VBProjec name is set
' according to the template

    On Error GoTo Test_vtkRecreateConfiguration_templateXlamVBProject_Error

    Set testedConf = xlamTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the VBProjectName
    mAssert.Equals recreatedWb.VBProject.name, "xlamTemplateProject"

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlamVBProject_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlamVBProject"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkRecreateConfiguration_templateXlamModule()
' ensure the configuration is well created with an xlam template, and that the module of the template
' exists in the recreated Workbook

    On Error GoTo Test_vtkRecreateConfiguration_templateXlammodule_Error

    Set testedConf = xlamTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the presence of the module from the template
    mAssert.Should Not recreatedWb.VBProject.VBComponents("xlamTemplateModule") Is Nothing

    mAssert.Should True, "The above shouldn't have thrown any error."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlammodule_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlammodule"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub Test_vtkRecreateConfiguration_templateXlamWorksheet()
' ensure the configuration is well created with an xlam template, and that the worksheet present in the template
' are present too in the recreatedWb

    On Error GoTo Test_vtkRecreateConfiguration_templateXlamWorksheet_Error

    Set testedConf = xlamTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the presence of the worksheet from the template
    mAssert.Should Not recreatedWb.Worksheets("xlamTemplateSheet") Is Nothing

    mAssert.Should True, "The above shouldn't have thrown any error."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlamWorksheet_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlamWorksheet"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub Test_vtkRecreateConfiguration_templateXlaVBProject()
' ensure the configuration is well created with an xla template, and that the VBProjec name is set
' according to the template

    On Error GoTo Test_vtkRecreateConfiguration_templateXlaVBProject_Error

    Set testedConf = xlaTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the VBProjectName
    mAssert.Equals recreatedWb.VBProject.name, "xlaTemplateProject"

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlaVBProject_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlaVBProject"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkRecreateConfiguration_templateXlaModule()
' ensure the configuration is well created with an xla template, and that the module of the template
' exists in the recreated Workbook

    On Error GoTo Test_vtkRecreateConfiguration_templateXlamodule_Error

    Set testedConf = xlaTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the presence of the module from the template
    mAssert.Should Not recreatedWb.VBProject.VBComponents("xlaTemplateModule") Is Nothing

    mAssert.Should True, "The above shouldn't have thrown any error."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlamodule_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlamodule"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub Test_vtkRecreateConfiguration_templateXlaWorksheet()
' ensure the configuration is well created with an xla template, and that the worksheet present in the template
' are present too in the recreatedWb

    On Error GoTo Test_vtkRecreateConfiguration_templateXlaWorksheet_Error

    Set testedConf = xlaTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the presence of the worksheet from the template
    mAssert.Should Not recreatedWb.Worksheets("xlaTemplateSheet") Is Nothing

    mAssert.Should True, "The above shouldn't have thrown any error."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlaWorksheet_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlaWorksheet"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkRecreateConfiguration_templateXlsVBProject()
' ensure the configuration is well created with an xls template, and that the VBProjec name is set
' according to the template

    On Error GoTo Test_vtkRecreateConfiguration_templateXlsVBProject_Error

    Set testedConf = xlsTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the VBProjectName
    mAssert.Equals recreatedWb.VBProject.name, "xlsTemplateProject"

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlsVBProject_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlsVBProject"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub Test_vtkRecreateConfiguration_templateXlsModule()
' ensure the configuration is well created with an xls template, and that the module of the template
' exists in the recreated Workbook

    On Error GoTo Test_vtkRecreateConfiguration_templateXlsmodule_Error

    Set testedConf = xlsTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the presence of the module from the template
    mAssert.Should Not recreatedWb.VBProject.VBComponents("xlsTemplateModule") Is Nothing

    mAssert.Should True, "The above shouldn't have thrown any error."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlsmodule_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlsmodule"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub Test_vtkRecreateConfiguration_templateXlsWorksheet()
' ensure the configuration is well created with an xls template, and that the worksheet present in the template
' are present too in the recreatedWb

    On Error GoTo Test_vtkRecreateConfiguration_templateXlsWorksheet_Error

    Set testedConf = xlsTemplateConf

    ' Create and open recreatedWb
    getModulesAndRecreateWb testProject, testedConf

    ' Test the presence of the worksheet from the template
    mAssert.Should Not recreatedWb.Worksheets("xlsTemplateSheet") Is Nothing

    mAssert.Should True, "The above shouldn't have thrown any error."

    On Error GoTo 0
    Exit Sub

Test_vtkRecreateConfiguration_templateXlsWorksheet_Error:
    Err.Source = "vtkRecreateConfigurationTester::Test_vtkRecreateConfiguration_templateXlsWorksheet"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_modulesNotYetExported_normalConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_modulesNotYetExported_emptyConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_pathDoesNotExistYet1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_pathDoesNotExistYet2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_allModulesArePresent_normalConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_referencesAreActivated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_badReference"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLSM"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLAM"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLS"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_formatLikeExtension_XLA"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookOpen"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_addInOpen_XLAM"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_workbookAndVbProjectAlreadyOpen"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_overwriteFile_normal"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_overwriteFile_addIn"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_title"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_comment"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlsmVBProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlsmModule"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlsmWorksheet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlamVBProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlamModule"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlamWorksheet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlaVBProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlaModule"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlaWorksheet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlsVBProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlsModule"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_templateXlsWorksheet"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_normalConf
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_emptyConf
        Case "Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf": Test_vtkRecreateConfiguration_workbookIsCreatedInTheGoodPath_multitypeConf
        Case "Test_vtkRecreateConfiguration_modulesNotYetExported_normalConf": Test_vtkRecreateConfiguration_modulesNotYetExported_normalConf
        Case "Test_vtkRecreateConfiguration_modulesNotYetExported_emptyConf": Test_vtkRecreateConfiguration_modulesNotYetExported_emptyConf
        Case "Test_vtkRecreateConfiguration_pathDoesNotExistYet1": Test_vtkRecreateConfiguration_pathDoesNotExistYet1
        Case "Test_vtkRecreateConfiguration_pathDoesNotExistYet2": Test_vtkRecreateConfiguration_pathDoesNotExistYet2
        Case "Test_vtkRecreateConfiguration_allModulesArePresent_normalConf": Test_vtkRecreateConfiguration_allModulesArePresent_normalConf
        Case "Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf": Test_vtkRecreateConfiguration_allModulesArePresent_multitypeConf
        Case "Test_vtkRecreateConfiguration_referencesAreActivated": Test_vtkRecreateConfiguration_referencesAreActivated
        Case "Test_vtkRecreateConfiguration_badReference": Test_vtkRecreateConfiguration_badReference
        Case "Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect": Test_vtkRecreateConfiguration_nameOfTheVBProjectIsCorrect
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLSM": Test_vtkRecreateConfiguration_formatLikeExtension_XLSM
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLAM": Test_vtkRecreateConfiguration_formatLikeExtension_XLAM
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLS": Test_vtkRecreateConfiguration_formatLikeExtension_XLS
        Case "Test_vtkRecreateConfiguration_formatLikeExtension_XLA": Test_vtkRecreateConfiguration_formatLikeExtension_XLA
        Case "Test_vtkRecreateConfiguration_workbookOpen": Test_vtkRecreateConfiguration_workbookOpen
        Case "Test_vtkRecreateConfiguration_addInOpen_XLAM": Test_vtkRecreateConfiguration_addInOpen_XLAM
        Case "Test_vtkRecreateConfiguration_workbookAndVbProjectAlreadyOpen": Test_vtkRecreateConfiguration_workbookAndVbProjectAlreadyOpen
        Case "Test_vtkRecreateConfiguration_overwriteFile_normal": Test_vtkRecreateConfiguration_overwriteFile_normal
        Case "Test_vtkRecreateConfiguration_overwriteFile_addIn": Test_vtkRecreateConfiguration_overwriteFile_addIn
        Case "Test_vtkRecreateConfiguration_title": Test_vtkRecreateConfiguration_title
        Case "Test_vtkRecreateConfiguration_comment": Test_vtkRecreateConfiguration_comment
        Case "Test_vtkRecreateConfiguration_templateXlsmVBProject": Test_vtkRecreateConfiguration_templateXlsmVBProject
        Case "Test_vtkRecreateConfiguration_templateXlsmModule": Test_vtkRecreateConfiguration_templateXlsmModule
        Case "Test_vtkRecreateConfiguration_templateXlsmWorksheet": Test_vtkRecreateConfiguration_templateXlsmWorksheet
        Case "Test_vtkRecreateConfiguration_templateXlamVBProject": Test_vtkRecreateConfiguration_templateXlamVBProject
        Case "Test_vtkRecreateConfiguration_templateXlamModule": Test_vtkRecreateConfiguration_templateXlamModule
        Case "Test_vtkRecreateConfiguration_templateXlamWorksheet": Test_vtkRecreateConfiguration_templateXlamWorksheet
        Case "Test_vtkRecreateConfiguration_templateXlaVBProject": Test_vtkRecreateConfiguration_templateXlaVBProject
        Case "Test_vtkRecreateConfiguration_templateXlaModule": Test_vtkRecreateConfiguration_templateXlaModule
        Case "Test_vtkRecreateConfiguration_templateXlaWorksheet": Test_vtkRecreateConfiguration_templateXlaWorksheet
        Case "Test_vtkRecreateConfiguration_templateXlsVBProject": Test_vtkRecreateConfiguration_templateXlsVBProject
        Case "Test_vtkRecreateConfiguration_templateXlsModule": Test_vtkRecreateConfiguration_templateXlsModule
        Case "Test_vtkRecreateConfiguration_templateXlsWorksheet": Test_vtkRecreateConfiguration_templateXlsWorksheet
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

