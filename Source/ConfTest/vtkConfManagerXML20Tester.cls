VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfManagerXML20Tester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkConfManagerXML20Tester
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Test the vtkConfigurationManagerXML class
'             with vtkConfigurations XML version 2.0
'
' Copyright 2014 Skwal-Soft (http://skwalsoft.com)
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

Private Const existingXMLNameForTest As String = "XMLForConfigurationsTests.xml"
Private existingConfManager As vtkConfigurationManager   ' Configuration Manager for the existing workbook
Private Const existingProjectName As String = "ExistingProject"

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
    Dim cmE As vtkConfigurationManagerXML, testFilePath As String, testFileName As String
    Set mAssert = Assert
    
    testFileName = existingProjectName & ".xml"
    testFilePath = VBAToolKit.vtkTestPath & "\" & testFileName
    getTestFileFromTemplate fileName:=existingXMLNameForTest, destinationName:=testFileName, openExcel:=False
    Set existingConfManager = New vtkConfigurationManagerXML
    Set cmE = existingConfManager
    cmE.init testFilePath
End Sub

Private Sub ITestCase_TearDown()
    VBAToolKit.resetTestFolder
End Sub

Public Sub Test_Init_FileNotFound()
    ' Verify that the proper error is raised in case of file not found
    Dim cm As New vtkConfigurationManagerXML
   On Error Resume Next
    cm.init "C:\NoFolder\NoFile.xml"
    mAssert.Equals Err.Number, VTK_WRONG_FILE_PATH, "Error returned when init the XML conf manager"
   On Error GoTo 0
End Sub

Public Sub Test_Init_BadXMLFile()
    ' Verify that the proper error is raised in case of bad XML File
    Dim cm As New vtkConfigurationManagerXML, filePath As String
    filePath = VBAToolKit.vtkTestPath & "\EmptyXMLFile.xml"
    Dim fso As New FileSystemObject
    fso.CreateTextFile filePath
   On Error Resume Next
    cm.init filePath
    mAssert.Equals Err.Number, VTK_INVALID_XML_FILE, "Error returned when init the XML conf manager"
   On Error GoTo 0
End Sub

Public Sub Test_Init_ObsoleteXMLFile()
    ' Verify that the proper error is raised in case of obsolete XML File
    Dim fileName As String, filePath As String, cm As New vtkConfigurationManagerXML
    fileName = "ExistingProject.xml"
    filePath = VBAToolKit.vtkTestPath & "\" & fileName
    getTestFileFromTemplate fileName:="EmptyXMLForConfigurationsTests.xml", destinationName:=fileName, openExcel:=False
   On Error Resume Next
    cm.init filePath
    mAssert.Equals Err.Number, VTK_OBSOLETE_CONFIGURATION_SHEET, "Error returned when init the XML conf manager"
   On Error GoTo 0
End Sub

Public Sub Test_ForbiddenSetterCalls()
    ' Verify that a call to a setter raise the proper error
   On Error Resume Next
    existingConfManager.addConfiguration "NewConfiguration", "ConfigurationPath"
    mAssert.Equals Err.Number, VTK_READONLY_FILE, "Error returned when trying to call addConfiguration on XML Conf File"
    Err.Number = 0
    existingConfManager.setConfigurationPathWithNumber n:=0, path:="Path0"
    mAssert.Equals Err.Number, VTK_READONLY_FILE, "Error returned when trying to call setConfigurationPathWithNumber on XML Conf File"
    Err.Number = 0
    existingConfManager.setConfigurationPath configuration:="InexistantConfiguration", path:="Path0"
    mAssert.Equals Err.Number, VTK_READONLY_FILE, "Error returned when trying to call setConfigurationPath on XML Conf File"
    Err.Number = 0
    existingConfManager.setConfigurationProjectNameWithNumber n:=1, projectName:="NewName"
    mAssert.Equals Err.Number, VTK_READONLY_FILE, "Error returned when trying to call setConfigurationProjectNameWithNumber on XML Conf File"
    Err.Number = 0
    existingConfManager.setConfigurationCommentWithNumber n:=1, comment:="NewComment"
    mAssert.Equals Err.Number, VTK_READONLY_FILE, "Error returned when trying to call setConfigurationCommentWithNumber on XML Conf File"
    Err.Number = 0
    existingConfManager.setConfigurationTemplateWithNumber n:=1, template:="NewTemplate"
    mAssert.Equals Err.Number, VTK_READONLY_FILE, "Error returned when trying to call setConfigurationTemplateWithNumber on XML Conf File"
    Err.Number = 0
    existingConfManager.addModule module:="NewModule1"
    mAssert.Equals Err.Number, VTK_READONLY_FILE, "Error returned when trying to call addModule on XML Conf File"
    Err.Number = 0
    existingConfManager.setModulePathWithNumber path:="NewPath", numModule:=1, numConfiguration:=1
    mAssert.Equals Err.Number, VTK_READONLY_FILE, "Error returned when trying to call setModulePathWithNumber on XML Conf File"
    Err.Number = 0
   On Error GoTo 0
End Sub

Public Sub Test_PropertyName_DefaultGet()
    '   Verify that the Property Name is the Default property for vtkConfigurationManager
    '   - In fact there is no need to run the test, just to compile it
    mAssert.Equals existingConfManager, "ExistingProject", "The name property must be the default one for vtkConfigurationManager"
End Sub

Public Sub TestGetConfigurationsFromExistingProject()
'       Verify the list of the configurations of an existing project
    mAssert.Equals existingConfManager.configurationCount, 2, "There must be two configurations in the existing template project"
    mAssert.Equals existingConfManager.configuration(0), "", "Inexistant configuration number 0"
    mAssert.Equals existingConfManager.configuration(1), existingProjectName, "Name of the first configuration"
    mAssert.Equals existingConfManager.configuration(2), existingProjectName & "_DEV", "Name of the second configuration"
    mAssert.Equals existingConfManager.configuration(3), "", "Inexistant configuration number 3"
    mAssert.Equals existingConfManager.configuration(-23), "", "Inexistant configuration number -23"
    mAssert.Equals existingConfManager.configuration(150), "", "Inexistant configuration number 150"
End Sub

Public Sub TestGetConfigurationPathWithNumberFromExistingProject()
'       Verify the capability to get the configuration path by number
    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(0), "", "Inexistant configuration number 0"
    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(1), "Delivery\ExistingProject.xlsm", "Path of first configuration given by number"
    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(2), "Project\ExistingProject_DEV.xlsm", "Path of second configuration given by number"
    mAssert.Equals existingConfManager.getConfigurationPathWithNumber(3), "", "Inexistant configuration number 3"
End Sub

Public Sub TestGetConfigurationProjectNameWithNumberFromExistingProject()
'       Verify the capability to get the configuration project name by number
    mAssert.Equals existingConfManager.getConfigurationProjectNameWithNumber(0), "", "Inexistant configuration number 0"
    mAssert.Equals existingConfManager.getConfigurationProjectNameWithNumber(1), "ExistingProjectName", "Project Name of first configuration given by number"
    mAssert.Equals existingConfManager.getConfigurationProjectNameWithNumber(2), "ExistingProject_DEV", "Project Name of second configuration given by number"
    mAssert.Equals existingConfManager.getConfigurationProjectNameWithNumber(3), "", "Inexistant configuration number 3"
End Sub

Public Sub TestGetConfigurationCommentWithNumberFromExistingProject()
'       Verify the capability to get the configuration comment by number
    mAssert.Equals existingConfManager.getConfigurationCommentWithNumber(0), "", "Inexistant configuration number 0"
    mAssert.Equals existingConfManager.getConfigurationCommentWithNumber(1), "Existing project for various tests of VBAToolKit", "Comment of first configuration given by number"
    mAssert.Equals existingConfManager.getConfigurationCommentWithNumber(2), "Existing project for development for various tests of VBAToolKit", "Comment of second configuration given by number"
    mAssert.Equals existingConfManager.getConfigurationCommentWithNumber(3), "", "Inexistant configuration number 3"
End Sub

Public Sub TestGetConfigurationTemplateWithNumberFromExistingProject()
'       Verify the capability to get the configuration template path by number
    mAssert.Equals existingConfManager.getConfigurationTemplateWithNumber(0), "", "Inexistant configuration number 0"
    mAssert.Equals existingConfManager.getConfigurationTemplateWithNumber(1), "", "Template path of first configuration given by number"
    mAssert.Equals existingConfManager.getConfigurationTemplateWithNumber(2), "Templates\ExistingProjectTemplate.xlsm", "Template path of second configuration given by number"
    mAssert.Equals existingConfManager.getConfigurationTemplateWithNumber(3), "", "Inexistant configuration number 3"
End Sub

Public Sub TestGetConfigurationNumbersFromExistingProject()
'       Verify the capability to get the number of a configuration
    mAssert.Equals existingConfManager.configurationCount, 2, "There must be two configurations in a new project"
    mAssert.Equals existingConfManager.getConfigurationNumber(existingProjectName), 1, "Number of the main configuration"
    mAssert.Equals existingConfManager.getConfigurationNumber(existingProjectName & "_DEV"), 2, "Number of the Development configuration"
    mAssert.Equals existingConfManager.getConfigurationNumber("InexistantConfiguration"), 0, "Inexistant configuration"
End Sub

Public Sub TestGetConfigurationPathFromExistingProject()
'       Verify the capability to get a configutaion path given the configuration name
    mAssert.Equals existingConfManager.getConfigurationPath(existingProjectName), "Delivery\ExistingProject.xlsm", "Path of the main configuration"
    mAssert.Equals existingConfManager.getConfigurationPath(existingProjectName & "_DEV"), "Project\ExistingProject_DEV.xlsm", "Path of the Development configuration"
    mAssert.Equals existingConfManager.getConfigurationPath("InexistantConfiguration"), "", "Inexistant configuration"
End Sub

Public Sub TestGetModulesFromExistingProject()
'       Verify the capability to retrieve the list of Modules from an existing project
    mAssert.Equals existingConfManager.moduleCount, 5, "There must be five configurations in the existing project"
    mAssert.Equals existingConfManager.module(0), "", "Inexistant module number 0"
    mAssert.Equals existingConfManager.module(1), "Module1", "Name of the first module"
    mAssert.Equals existingConfManager.module(2), "Module2", "Name of the second module"
    mAssert.Equals existingConfManager.module(3), "Module3", "Name of the third module"
    mAssert.Equals existingConfManager.module(4), "Module4", "Name of the fourth module"
    mAssert.Equals existingConfManager.module(5), "Module5", "Name of the fifth module"
    mAssert.Equals existingConfManager.module(6), "", "Inexistant module number 6"
    mAssert.Equals existingConfManager.module(-23), "", "Inexistant module number -23"
    mAssert.Equals existingConfManager.module(150), "", "Inexistant module number 150"
End Sub

Public Sub TestGetModuleNumbersFromExistingProject()
'       Verify the capability to get the number of a configuration
    mAssert.Equals existingConfManager.getModuleNumber("Module0"), 0, "Inexistant module"
    mAssert.Equals existingConfManager.getModuleNumber("Module1"), 1, "First Module"
    mAssert.Equals existingConfManager.getModuleNumber("Module2"), 2, "Second Module"
    mAssert.Equals existingConfManager.getModuleNumber("Module3"), 3, "Third module"
    mAssert.Equals existingConfManager.getModuleNumber("Module4"), 4, "Fourth module"
    mAssert.Equals existingConfManager.getModuleNumber("Module5"), 5, "Fifth module"
    mAssert.Equals existingConfManager.getModuleNumber("InexistantModule"), 0, "Inexistant module"
End Sub

Public Sub TestGetModulePathWithNumberFromExistingProject()
'       Verify the capability to get the module path by number
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=0, numConfiguration:=2), "", "Inexistant module path number 0,2"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=3), "", "Inexistant module path number 3,3"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=1), "Path1Module1", "Module path number 1,1"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=1, numConfiguration:=2), "", "Module path number 1,2"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=2, numConfiguration:=1), "", "Module path number 2,1"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=2, numConfiguration:=2), "Path2Module2", "Module path number 2,2"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=1), "", "Module path number 3,1"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=3, numConfiguration:=2), "", "Module path number 3,2"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=4, numConfiguration:=1), "Path1Module4", "Module path number 4,1"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=4, numConfiguration:=2), "Path2Module4", "Module path number 4,2"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=5, numConfiguration:=1), "", "Module path number 5,1"
    mAssert.Equals existingConfManager.getModulePathWithNumber(numModule:=5, numConfiguration:=2), "Path2Module5", "Module path number 5,2"
End Sub

Public Sub TestModulesFromExistingProject()
'       Verify the capability to get the modules list
    Dim n As Integer
   On Error GoTo M_Error
    mAssert.Equals existingConfManager.modules.count, 5, "Number of modules in the existing project"
    For n = 1 To n
        mAssert.Should existingConfManager.modules(n).parent Is existingConfManager, "Parent of the vtkModule " & n & " of existing project"
        mAssert.Equals existingConfManager.modules(n).Number, n, "The number of vtkModule " & n & " of existing project"
        mAssert.Equals existingConfManager.modules(n).id, "m" & n, "Expected ID for module " & n
        mAssert.Equals existingConfManager.modules(n).name, "Module" & n, "Name of the module " & n & " get by number"
        mAssert.Equals existingConfManager.modules("Module" & n).Number, n, "Number of the module " & n & " get by name"
    Next n
   On Error GoTo 0
    Exit Sub
M_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description
End Sub

Public Sub TestConfigurationsCollectionFromExistingProject()
'       Verify the behavior of the configurations function of Conf manager
   On Error GoTo M_Error
    mAssert.Equals existingConfManager.configurations.count, 2, "Number of configurations in the existing project"
    ' Check the first configuration
    mAssert.Should existingConfManager.configurations(1).parent Is existingConfManager, "Parent of the first configuration"
    mAssert.Equals existingConfManager.configurations(1).Number, 1, "Number of the first configuration"
    mAssert.Equals existingConfManager.configurations(1).id, "c1", "Expected ID for the first vtkConfiguration"
    mAssert.Should Not existingConfManager.configurations(1).isDEV, "The first configuration is expected not to be DEV"
    mAssert.Equals existingConfManager.configurations(1).name, "ExistingProject", "Name of the configuration for the first vtkConfiguration"
    mAssert.Equals existingConfManager.configurations("ExistingProject").name, "ExistingProject", "Name of the configuration for the first vtkConfiguration"
    mAssert.Equals existingConfManager.configurations(1).comment, "Existing project for various tests of VBAToolKit", "Comment got by number of the first configuration"
    mAssert.Equals existingConfManager.configurations("ExistingProject").comment, "Existing project for various tests of VBAToolKit", "Comment got by name of the first configuration"
    mAssert.Equals existingConfManager.configurations(1).template, "", "Template got by number of the first configuration"
    mAssert.Equals existingConfManager.configurations("ExistingProject").template, "", "Template got by name of the first configuration"
    mAssert.Equals existingConfManager.configurations(1).projectName, "ExistingProjectName", "Project Name (Title) got by number of the first configuration"
    mAssert.Equals existingConfManager.configurations("ExistingProject").projectName, "ExistingProjectName", "Project Name (Title) got by name of the first configuration"
    ' Check the second configuration
    mAssert.Should existingConfManager.configurations(2).parent Is existingConfManager, "Parent of the second configuration"
    mAssert.Equals existingConfManager.configurations(2).Number, 2, "Number of the second configuration"
    mAssert.Equals existingConfManager.configurations(2).id, "c2", "Expected ID for the second vtkConfiguration"
    mAssert.Should existingConfManager.configurations(2).isDEV, "The second configuration is expected to be DEV"
    mAssert.Equals existingConfManager.configurations(2).name, "ExistingProject_DEV", "Name of the configuration for the second vtkConfiguration"
    mAssert.Equals existingConfManager.configurations("ExistingProject_DEV").name, "ExistingProject_DEV", "Name of the configuration for the second vtkConfiguration"
    mAssert.Equals existingConfManager.configurations(2).comment, "Existing project for development for various tests of VBAToolKit", "Comment got by number of the second configuration"
    mAssert.Equals existingConfManager.configurations("ExistingProject_DEV").comment, "Existing project for development for various tests of VBAToolKit", "Comment got by name of the second configuration"
    mAssert.Equals existingConfManager.configurations(2).template, "Templates\ExistingProjectTemplate.xlsm", "Template got by number of the second configuration"
    mAssert.Equals existingConfManager.configurations("ExistingProject_DEV").template, "Templates\ExistingProjectTemplate.xlsm", "Template got by name of the second configuration"
    mAssert.Equals existingConfManager.configurations(2).projectName, "ExistingProject_DEV", "Project Name (Title) got by number of the second configuration"
    mAssert.Equals existingConfManager.configurations("ExistingProject_DEV").projectName, "ExistingProject_DEV", "Project Name (Title) got by name of the second configuration"
   On Error GoTo 0
    Exit Sub
M_Error:
    mAssert.Should False, "Unexpected error (" & Err.Number & ") : " & Err.Description
End Sub

Public Sub TestGetConfigurationReferencesWithNumber()
'       Check the behavior of the getConfigurationReferencesWithNumber function
    Dim refNames(), i As Integer, cn As Integer, c1 As Collection, c2 As Collection, r As vtkReference
    Dim expectedCount As Integer
    refNames = Array("VBA", "Excel", "stdole", "Office", "MSForms", "Scripting", "VBIDE", "Shell32", "MSXML2")
   On Error Resume Next
    For cn = 1 To existingConfManager.configurationCount
        Set c1 = existingConfManager.configurations(cn).references
        ' Rearrange the collection by name
        Set c2 = New Collection
        For Each r In c1
            c2.Add r, r.name
        Next
        expectedCount = UBound(refNames) - LBound(refNames) + 1
        If existingConfManager.configurations(cn).isDEV Then expectedCount = expectedCount + 1
        mAssert.Equals c2.count, expectedCount, "Count of references in configuration " & cn
        ' il faut boucler sur le tableau et rechercher dans la collection (si pas trouvé = erreur)
        For i = LBound(refNames) To UBound(refNames)
            Set r = c2(refNames(i))
            mAssert.Equals Err.Number, 0, "Error when getting " & refNames(i) & " reference from configuration " & cn
        Next i
        If existingConfManager.configurations(cn).isDEV Then
            Set r = c2("VBAToolKit")
            mAssert.Equals Err.Number, 0, "Error when getting VBAToolKit reference from configuration " & cn
        End If
    Next cn
   On Error GoTo 0
End Sub

Public Sub TestRootPathForExistingProject()
    mAssert.Equals existingConfManager.rootPath, vtkPathOfCurrentProject, "The root Path is not initialized for a new Workbook"
    mAssert.Equals existingConfManager.rootPath, vtkPathOfCurrentProject, "The second call to rootPath give the same result as the previous one"
End Sub

Public Sub TestGetAllReferencesFromExistingWorkbook()
'       Verify that all references are listed from existing workbook
    Dim refNames(), i As Integer, c1 As Collection, c2 As Collection, r As vtkReference
    refNames = Array("VBA", "Excel", "stdole", "Office", "MSForms", "Scripting", "VBIDE", "Shell32", "MSXML2", "VBAToolKit", "EventSystemLib")

   On Error Resume Next
    Set c1 = existingConfManager.references
    ' Rearrange the collection by name
    Set c2 = New Collection
    For Each r In c1
        c2.Add r, r.name
    Next
    mAssert.Equals c2.count, UBound(refNames) - LBound(refNames) + 1, "Count of all references of a new workbook"
    ' il faut boucler sur le tableau et rechercher dans la collection (si pas trouvé = erreur)
    For i = LBound(refNames) To UBound(refNames)
        Set r = c2(refNames(i))
        mAssert.Equals Err.Number, 0, "Error when getting " & refNames(i) & " reference"
    Next i
   On Error GoTo 0
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_Init_FileNotFound"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_Init_BadXMLFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_Init_ObsoleteXMLFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ForbiddenSetterCalls"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_PropertyName_DefaultGet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationsFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationPathWithNumberFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationProjectNameWithNumberFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationCommentWithNumberFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationTemplateWithNumberFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationNumbersFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationPathFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetModulesFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetModuleNumbersFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetModulePathWithNumberFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestModulesFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfigurationsCollectionFromExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetConfigurationReferencesWithNumber"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestRootPathForExistingProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetAllReferencesFromExistingWorkbook"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_Init_FileNotFound": Test_Init_FileNotFound
        Case "Test_Init_BadXMLFile": Test_Init_BadXMLFile
        Case "Test_Init_ObsoleteXMLFile": Test_Init_ObsoleteXMLFile
        Case "Test_ForbiddenSetterCalls": Test_ForbiddenSetterCalls
        Case "Test_PropertyName_DefaultGet": Test_PropertyName_DefaultGet
        Case "TestGetConfigurationsFromExistingProject": TestGetConfigurationsFromExistingProject
        Case "TestGetConfigurationPathWithNumberFromExistingProject": TestGetConfigurationPathWithNumberFromExistingProject
        Case "TestGetConfigurationProjectNameWithNumberFromExistingProject": TestGetConfigurationProjectNameWithNumberFromExistingProject
        Case "TestGetConfigurationCommentWithNumberFromExistingProject": TestGetConfigurationCommentWithNumberFromExistingProject
        Case "TestGetConfigurationTemplateWithNumberFromExistingProject": TestGetConfigurationTemplateWithNumberFromExistingProject
        Case "TestGetConfigurationNumbersFromExistingProject": TestGetConfigurationNumbersFromExistingProject
        Case "TestGetConfigurationPathFromExistingProject": TestGetConfigurationPathFromExistingProject
        Case "TestGetModulesFromExistingProject": TestGetModulesFromExistingProject
        Case "TestGetModuleNumbersFromExistingProject": TestGetModuleNumbersFromExistingProject
        Case "TestGetModulePathWithNumberFromExistingProject": TestGetModulePathWithNumberFromExistingProject
        Case "TestModulesFromExistingProject": TestModulesFromExistingProject
        Case "TestConfigurationsCollectionFromExistingProject": TestConfigurationsCollectionFromExistingProject
        Case "TestGetConfigurationReferencesWithNumber": TestGetConfigurationReferencesWithNumber
        Case "TestRootPathForExistingProject": TestRootPathForExistingProject
        Case "TestGetAllReferencesFromExistingWorkbook": TestGetAllReferencesFromExistingWorkbook
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


