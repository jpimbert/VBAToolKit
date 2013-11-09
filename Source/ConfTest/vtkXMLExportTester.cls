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
Private conf As vtkConfiguration
Private testFilePath As String

Private retDom As MSXML2.DOMDocument
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

    Set retDom = vtkExportAsXMLDOM(testProjectName)
    
End Sub


Private Sub ITestCase_TearDown()
    On Error Resume Next
    
    vtkResetConfigurationManagers
    vtkCloseAndKillWorkbook Wb:=testWorkbook
    VBAToolKit.resetTestFolder
    
End Sub



' -------------------------------------------
' TESTS TO DO
'
' exportDom
'   rootNode -> correct name
'   info -> 1 info object
'   info -> correct content
'   conf -> correct number of configurations
'   conf -> correct content
'   modules -> correct number of modules
'   modules -> correct content
'   dummy project name -> raise VTK_PROJECT_NOT_INITIALIZED
'
' WriteDom
'   dummy DOM -> raise error VTK_DOM_NOT_INITIALIZED
'   file does not yet exist -> file is created
'   file already exists -> file is overwritten
'   file path does not exist -> raise error VTK_WRONG_FILE_PATH
'   load content of the file -> file content is a valid DOM
'
'--------------------------------------------

'--------------------------------------------
' Utility functions
'--------------------------------------------
Private Function countElementsInDom(elementName As String, dom As MSXML2.DOMDocument) As Integer

    On Error GoTo countElementsInDom_Error
    
    Dim rootNode As MSXML2.IXMLDOMNode
    Set rootNode = dom.ChildNodes.Item(1)
    
    countElementsInDom = countElementsInNode(elementName, rootNode)

    On Error GoTo 0
    Exit Function

countElementsInDom_Error:
    Debug.Print "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure countElementsInDom of Class Module vtkXMLExportTester"
    Exit Function
    
End Function

Private Function countElementsInNode(elementName As String, node As MSXML2.IXMLDOMNode) As Integer
    
    Dim Count As Integer: Count = 0
    
    On Error GoTo countElementsInNode_Error

    Dim subNode As MSXML2.IXMLDOMNode
    For Each subNode In node.ChildNodes
        If StrComp(subNode.BaseName, elementName) = 0 Then Count = Count + 1
    Next
        
    countElementsInNode = Count

    On Error GoTo 0
    Exit Function

countElementsInNode_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure countElementsInNode of Class Module vtkXMLExportTester"
    ITestCase_TearDown
    Exit Function
End Function

'--------------------------------------------
' Tests
'--------------------------------------------

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


Public Sub Test_vtkExportAsXMLDom_dummyProject()
'   dummy project name -> raise VTK_PROJECT_NOT_INITIALIZED

    On Error GoTo Test_vtkExportAsXMLDom_dummyProject_Error
    
    Dim error As Integer: error = 0
    Dim dummyProjectName As String
    dummyProjectName = "dummyProject"
    
    Set retDom = vtkExportAsXMLDOM(dummyProjectName)
    
    mAssert.Equals error, VTK_PROJECT_NOT_INITIALIZED, "Project is empty : VTK_PROJECT_NOT_INITIALIZED should be raised."

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_dummyProject_Error:
    error = Err.Number
    Resume Next
End Sub


Public Sub Test_vtkWriteXMLDomToFile_dummyDom()
'   dummy DOM -> raise error VTK_DOM_NOT_INITIALIZED, no file should be created

    On Error GoTo Test_vtkWriteXMLDomToFile_dummyDom_Error

    Dim error As Integer
    Dim dummyDom As MSXML2.DOMDocument

    fso.DeleteFile testFilePath
    vtkWriteXMLDOMToFile dummyDom, testFilePath
    
    mAssert.Equals error, VTK_DOM_NOT_INITIALIZED, "Dom is not initialized : VTK_DOM_NOT_INITIALIZED should be raised."
    
    mAssert.Should (fso.FileExists(testFilePath) = False), "File should not have been created as there has been an error."
    
    On Error GoTo 0
    Exit Sub

Test_vtkWriteXMLDomToFile_dummyDom_Error:
    error = Err.Number
    Resume Next
End Sub


Public Sub Test_vtkWriteXMLDomToFile_fileIsCreated()
'   file does not exist -> file is created

    On Error GoTo Test_vtkvtkWriteXMLDomToFile_fileIsCreated_Error

    Dim fso As New FileSystemObject

    vtkWriteXMLDOMToFile retDom, testFilePath
    
    mAssert.Should fso.FileExists(testFilePath), "File should have been created as everything is normal."

    On Error GoTo 0
    Exit Sub

Test_vtkvtkWriteXMLDomToFile_fileIsCreated_Error:
    Err.Source = "Test_vtkvtkWriteXMLDomToFile_fileIsCreated of module aavtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub


Public Sub Test_vtkWriteXMLDomToFile_fileIsOverWritten()
'   file already exists -> file is overwritten

    On Error GoTo Test_vtkWriteXMLDomToFile_fileIsOverWritten_Error

    Dim fso As New FileSystemObject
    
    Dim oldSize As Integer
    Dim newSize As Integer
    
    fso.CreateTextFile testFilePath
    oldSize = fso.GetFile(testFilePath).Size

    vtkWriteXMLDOMToFile retDom, testFilePath
    
    newSize = fso.GetFile(testFilePath).Size
    
    mAssert.Should oldSize < newSize, "File should have been overwritten."
    
    On Error GoTo 0
    Exit Sub

Test_vtkWriteXMLDomToFile_fileIsOverWritten_Error:
    Err.Source = "Test_vtkWriteXMLDomToFile_fileIsOverWritten of module aavtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub


Public Sub Test_vtkWriteXMLDomToFile_filepathIsCreated()
'   file path does not exist -> raise error VTK_WRONG_FILE_PATH
    
    On Error GoTo Test_vtkWriteXMLDomTOFile_filepathIsCreated_Error
    
    Dim error As Integer: error = 0
    
    Dim newFilePath As String
    newFilePath = VBAToolKit.vtkTestPath & "\dummyFolder\testFile.xml"

    vtkWriteXMLDOMToFile retDom, newFilePath
    
    Dim fso As New FileSystemObject
    mAssert.Should Not fso.FileExists(newFilePath), "File should not have been created."

    mAssert.Equals error, VTK_WRONG_FILE_PATH

    On Error GoTo 0
    Exit Sub

Test_vtkWriteXMLDomTOFile_filepathIsCreated_Error:
    error = Err.Number
    Resume Next
End Sub


Public Sub Test_vtkWriteXMLDomToFile_fileContentIsValidDom()
'   load content of the file -> file content is a valid DOM

    On Error GoTo Test_vtkWriteXMLDomToFile_fileContentIsValidDom_Error
    
    vtkWriteXMLDOMToFile retDom, testFilePath
    
    Dim newDom As MSXML2.DOMDocument
    Set newDom = New MSXML2.DOMDocument
    newDom.Load testFilePath

    mAssert.Should countElementsInDom("info", newDom) = 1, "Dom should be readable."

    On Error GoTo 0
    Exit Sub

Test_vtkWriteXMLDomToFile_fileContentIsValidDom_Error:
    Err.Source = "Test_vtkWriteXMLDomToFile_fileContentIsValidDom of module aavtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
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
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_dummyProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkWriteXMLDomToFile_dummyDom"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkWriteXMLDomToFile_fileIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkWriteXMLDomToFile_fileIsOverWritten"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkWriteXMLDomToFile_filepathIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkWriteXMLDomToFile_fileContentIsValidDom"
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
        Case "Test_vtkExportAsXMLDom_dummyProject": Test_vtkExportAsXMLDom_dummyProject
        Case "Test_vtkWriteXMLDomToFile_dummyDom": Test_vtkWriteXMLDomToFile_dummyDom
        Case "Test_vtkWriteXMLDomToFile_fileIsCreated": Test_vtkWriteXMLDomToFile_fileIsCreated
        Case "Test_vtkWriteXMLDomToFile_fileIsOverWritten": Test_vtkWriteXMLDomToFile_fileIsOverWritten
        Case "Test_vtkWriteXMLDomToFile_filepathIsCreated": Test_vtkWriteXMLDomToFile_filepathIsCreated
        Case "Test_vtkWriteXMLDomToFile_fileContentIsValidDom": Test_vtkWriteXMLDomToFile_fileContentIsValidDom
        Case "Test_vtkExportConfigurationsAsXML_NoProject": Test_vtkExportConfigurationsAsXML_NoProject
        Case "Test_vtkExportConfigurationsAsXML_WrongPath": Test_vtkExportConfigurationsAsXML_WrongPath
        Case "Test_vtkExportConfigurationsAsXML_InexistantFile": Test_vtkExportConfigurationsAsXML_InexistantFile
        Case "Test_vtkExportConfigurationsAsXML_ExistantFile": Test_vtkExportConfigurationsAsXML_ExistantFile
        Case "Test_vtkExportConfigurationsAsXML_fileContentIsValidDom": Test_vtkExportConfigurationsAsXML_fileContentIsValidDom
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

