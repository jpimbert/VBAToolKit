VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkXMLExportTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkXMLExportTester
' Author    : Lucas Vitorino
' Purpose   : Test the vtkExportAsDOM and vtkWriteDOMToFile functions.
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


Private Const testProjectName As String = "TestProject"
Private Const testProjectFileName As String = testProjectName & "_DEV.xlsm"
Private Const testFileName As String = "testFile.xml"

Private fso As New FileSystemObject

Private testWorkbook As Workbook
Private cm As vtkConfigurationManager
Private testFilePath As String

Private reloadDom As MSXML2.DOMDocument

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
    
    ' Copy the DTD in the Test Folder
    getTestFileFromTemplate fileName:="vtkConfigurationsDTD.dtd", openExcel:=False
    
    ' Open the test Project and get the configuration manager for this project
    Set testWorkbook = getTestFileFromTemplate(fileName:=testProjectFileName, openExcel:=True)
    Set cm = vtkConfigurationManagerForProject(testProjectName)

    ' Export the test project configuration in XML format
    testFilePath = VBAToolKit.vtkTestPath & "\" & testFileName
    vtkExportConfigurationsAsXML projectName:=testProjectName, filePath:=testFilePath
    
    ' Load the configuration as a DOM structure for content testing
    Set reloadDom = New MSXML2.DOMDocument
    reloadDom.Load testFilePath
    If reloadDom.parseError.ErrorCode <> 0 Then
        Debug.Print "Parse error of " & testFilePath & " in SetUp of vtkXMLExportTester: " & reloadDom.parseError.reason
    End If
    
End Sub

Private Sub ITestCase_TearDown()
    On Error Resume Next
    
    vtkResetConfigurationManagers
    vtkCloseAndKillWorkbook Wb:=testWorkbook
    VBAToolKit.resetTestFolder
    
End Sub

Public Sub Test_vtkExportConfigurationsAsXML_correctInfoContent()
'   info -> correct content

    On Error GoTo Test_vtkExportConfigurationsAsXML_correctDomContent_Error
    
    mAssert.Equals reloadDom.SelectSingleNode("/vtkConf/info/projectName").Text, testProjectName, "Expected name of project in info element"
    mAssert.Equals reloadDom.SelectSingleNode("/vtkConf/info/vtkConfigurationsVersion").Text, "1.0", "Expected configuration version in info element"

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfigurationsAsXML_correctDomContent_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkExportConfigurationsAsXML_correctInfoContent of Class Module vtkXMLExportTester"
    Resume Next
End Sub


Public Sub Test_vtkExportConfigurationsAsXML_correctConfigurationNumber()
'   configurations -> correct number of configurations
    
    On Error GoTo Test_vtkExportConfigurationsAsXML_correctConfigurationNumber_Error
    
    mAssert.Equals reloadDom.SelectNodes("/vtkConf/configuration").Length, cm.configurationCount, "Expected number of configuration elements"

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfigurationsAsXML_correctConfigurationNumber_Error:
    Err.Source = "Test_vtkExportConfigurationsAsXML_correctConfigurationNumber of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Resume Next
End Sub


Public Sub Text_vtkExportConfigurationsAsXML_correctConfigurationContent()
'   configurations -> correct content

    On Error GoTo Text_vtkExportConfigurationsAsXML_correctConfigurationContent_Error

    mAssert.Equals reloadDom.SelectSingleNode("/vtkConf/configuration[@cID=""c1""]/name").Text, "TestProject", "Expected name of first configuration"
    mAssert.Equals reloadDom.SelectSingleNode("/vtkConf/configuration[@cID=""c1""]/path").Text, "Project\TestProject.xlsm", "Expected path of first configuration"
    
    mAssert.Equals reloadDom.SelectSingleNode("/vtkConf/configuration[@cID=""c2""]/name").Text, "TestProject_DEV", "Expected name of second configuration"
    mAssert.Equals reloadDom.SelectSingleNode("/vtkConf/configuration[@cID=""c2""]/path").Text, "Project\TestProject_DEV.xlsm", "Expected path of second configuration"

    On Error GoTo 0
    Exit Sub

Text_vtkExportConfigurationsAsXML_correctConfigurationContent_Error:
    Err.Source = "Text_vtkExportConfigurationsAsXML_correctConfigurationContent of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub


Public Sub Test_vtkExportConfigurationsAsXML_correctModuleNumber()
'   modules -> correct number of modules

    On Error GoTo Test_vtkExportConfigurationsAsXML_correctModuleNumber_Error
    
    mAssert.Equals reloadDom.SelectNodes("/vtkConf/module").Length, cm.moduleCount, "Expected number of module elements"

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfigurationsAsXML_correctModuleNumber_Error:
    Err.Source = "Test_vtkExportConfigurationsAsXML_correctModuleNumber of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub


Public Sub Test_vtkExportConfigurationsAsXML_correctModuleContent()
'   modules -> correct content
    
    On Error GoTo Test_vtkExportConfigurationsAsXML_correctModuleContent_Error
    
    mAssert.Equals reloadDom.SelectSingleNode("/vtkConf/module[@mID=""m1""]/name").Text, "TestModule1", "Expected name of first module"
    mAssert.Equals reloadDom.SelectSingleNode("/vtkConf/module[@mID=""m1""]/modulePath[@confId=""c2""]").Text, "Project\TestModule1.bas", "Expected path of first module for the second configuration"
    
    mAssert.Equals reloadDom.SelectSingleNode("/vtkConf/module[@mID=""m2""]/name").Text, "TestModule2", "Expected name of second module"
    mAssert.Equals reloadDom.SelectSingleNode("/vtkConf/module[@mID=""m2""]/modulePath[@confId=""c2""]").Text, "Project\TestModule2.bas", "Expected path of second module for the second configuration"

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfigurationsAsXML_correctModuleContent_Error:
    Err.Source = "Test_vtkExportConfigurationsAsXML_correctModuleContent of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub

Public Sub Test_vtkExportConfigurationsAsXML_NoProject()
'   Verifies that an attempt to export a nonexistant project configuration raises the proper error

   On Error GoTo Test_vtkExportConfigurationsAsXML_NoProject_Error

    vtkExportConfigurationsAsXML projectName:="InexistantProject", filePath:=""
    mAssert.Should False, "An error should occur when exporting an inexistant project"

   On Error GoTo 0
   Exit Sub

Test_vtkExportConfigurationsAsXML_NoProject_Error:
    mAssert.Equals Err.Number, VTK_WORKBOOK_NOT_OPEN, "Bad error raised when exporting an inexistant project"
End Sub

Public Sub Test_vtkExportConfigurationsAsXML_WrongPath()
'   Verifies that an attempt to export a project configuration in a bad path raises the proper error

   On Error GoTo Test_vtkExportConfigurationsAsXML_WrongPath_Error

    vtkExportConfigurationsAsXML projectName:=testProjectName, filePath:="c:\InexistantFolder\file.xml"
    mAssert.Should False, "An error should occur when exporting in a wrong path"

   On Error GoTo 0
   Exit Sub

Test_vtkExportConfigurationsAsXML_WrongPath_Error:
    mAssert.Equals Err.Number, VTK_WRONG_FILE_PATH, "Bad error raised when exporting in a wrong path"
End Sub

Public Sub Test_vtkExportConfigurationsAsXML_InexistantFile()
'   Verifies that a new configuration file is created if not yet existing

   On Error GoTo Test_vtkExportConfigurationsAsXML_InexistantFile_Error

    vtkExportConfigurationsAsXML projectName:=testProjectName, filePath:=testFilePath
    mAssert.Should fso.FileExists(testFilePath), "The configuration file must be created when exporting a project"

   On Error GoTo 0
   Exit Sub

Test_vtkExportConfigurationsAsXML_InexistantFile_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ")"
End Sub

Public Sub Test_vtkExportConfigurationsAsXML_ExistantFile()
'   Verifies that an old configuration file is erased and rewritten

   On Error GoTo Test_vtkExportConfigurationsAsXML_ExistantFile_Error

    ' Create an empty file
    Dim f As TextStream
    Set f = fso.CreateTextFile(testFilePath)
    f.WriteLine ("Unexpected line in a XML file")
    f.Close
    
    vtkExportConfigurationsAsXML projectName:=testProjectName, filePath:=testFilePath
    ' Verifies the first line
    Set f = fso.OpenTextFile(testFilePath)
    Dim line As String
    line = f.ReadLine
    f.Close
    mAssert.Equals line, "<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?>", "First line of the configuration file"

   On Error GoTo 0
   Exit Sub

Test_vtkExportConfigurationsAsXML_ExistantFile_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ")"
End Sub

Public Sub Test_vtkExportConfigurationsAsXML_fileContentIsValidDom()
'   load content of the file -> file content is a valid DOM

    On Error GoTo Test_vtkExportConfigurationsAsXML_fileContentIsValidDom_Error
    
    mAssert.Equals reloadDom.parseError.ErrorCode, 0, "Dom should be readable without parse error"

    On Error GoTo 0
    Exit Sub

Test_vtkExportConfigurationsAsXML_fileContentIsValidDom_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ")"
End Sub

'   Other tests to write
'
'   Verifies the content of the generated configuration file
'   - for these tests the XML file is created then re-read in a DOM structure
'   - the the old tests are recycled to work on the re-read structure rather than on an exported DOM structure

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfigurationsAsXML_correctInfoContent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfigurationsAsXML_correctConfigurationNumber"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfigurationsAsXML_correctModuleNumber"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfigurationsAsXML_correctModuleContent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfigurationsAsXML_NoProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfigurationsAsXML_WrongPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfigurationsAsXML_InexistantFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfigurationsAsXML_ExistantFile"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportConfigurationsAsXML_fileContentIsValidDom"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkExportConfigurationsAsXML_correctInfoContent": Test_vtkExportConfigurationsAsXML_correctInfoContent
        Case "Test_vtkExportConfigurationsAsXML_correctConfigurationNumber": Test_vtkExportConfigurationsAsXML_correctConfigurationNumber
        Case "Test_vtkExportConfigurationsAsXML_correctModuleNumber": Test_vtkExportConfigurationsAsXML_correctModuleNumber
        Case "Test_vtkExportConfigurationsAsXML_correctModuleContent": Test_vtkExportConfigurationsAsXML_correctModuleContent
        Case "Test_vtkExportConfigurationsAsXML_NoProject": Test_vtkExportConfigurationsAsXML_NoProject
        Case "Test_vtkExportConfigurationsAsXML_WrongPath": Test_vtkExportConfigurationsAsXML_WrongPath
        Case "Test_vtkExportConfigurationsAsXML_InexistantFile": Test_vtkExportConfigurationsAsXML_InexistantFile
        Case "Test_vtkExportConfigurationsAsXML_ExistantFile": Test_vtkExportConfigurationsAsXML_ExistantFile
        Case "Test_vtkExportConfigurationsAsXML_fileContentIsValidDom": Test_vtkExportConfigurationsAsXML_fileContentIsValidDom
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

