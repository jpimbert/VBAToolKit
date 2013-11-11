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
Private VTKConf As vtkConfiguration
Private VTKProjectConf As vtkConfiguration
Private testedConf As vtkConfiguration

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
Private Const VTKProjectConfName = "VBAToolKit"

Private Const standardRememberedProjectsDTDName As String = "rememberedProjectsDTD.dtd"
Private Const standardVtkConfigurationsDTDName As String = "vtkConfigurationsDTD.dtd"
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
    createDTDforRememberedProjects sheetpath:=fso.BuildPath(VBAToolKit.vtkTestPath, standardRememberedProjectsDTDName)
    
    createRememberedProjectsXMLSheet sheetpath:=fso.BuildPath(VBAToolKit.vtkTestPath, rememberedProjectsSheetName), _
                                     dtdPath:=fso.BuildPath(VBAToolKit.vtkTestPath, standardRememberedProjectsDTDName)
    
    ' Copy the sheet of the new project in the relevant folder, and create a DTD
    FileCopy Source:=fso.BuildPath(vtkPathToTemplateFolder, testProject.XMLConfigurationName), _
             destination:=fso.BuildPath(testProjectFolder, testProject.XMLConfigurationStandardRelativePath)
    
    createDTDForVtkConfigurations fso.BuildPath(testProjectFolder, "Project\" & standardVtkConfigurationsDTDName)
    
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
    Debug.Print "Error " & Err.Number & " : " & Err.Description & " in " & Err.Source
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
    Debug.Print "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub


' --------------------------------------------------------
'                       TESTS
' --------------------------------------------------------

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

    Set testedConf = VTKConf

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


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkRecreateConfiguration_referencesAreActivated"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkRecreateConfiguration_referencesAreActivated": Test_vtkRecreateConfiguration_referencesAreActivated
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

