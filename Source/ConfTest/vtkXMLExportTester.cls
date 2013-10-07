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

Private testWorkbook As Workbook
Private cm As vtkConfigurationManager
Private conf As vtkConfiguration
Private testFilePath As String

Private retDom As MSXML2.DOMDocument

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
    
    Set testWorkbook = getTestFileFromTemplate(fileName:=testProjectFileName, openExcel:=True)
    Set cm = vtkConfigurationManagerForProject(testProjectName)

    testFilePath = VBAToolKit.vtkTestPath & "\" & testFileName

    Set retDom = vtkExportAsXMLDOM(testProjectName)
    
End Sub


Private Sub ITestCase_TearDown()
    On Error Resume Next
    
    vtkResetConfigurationManagers
    vtkCloseAndKillWorkbook Wb:=testWorkbook
    vtkCleanFolder VBAToolKit.vtkTestPath
    
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

Public Sub Test_vtkExportAsXMLDom_correctNameForRoot()
'   rootNode -> correct name

    On Error GoTo Test_vtkExportAsXMLDom_correctNameForRoot_Error

    Dim expected As String
    Dim actual As String
    expected = "vtkConf"
    actual = retDom.ChildNodes.Item(1).BaseName

    mAssert.Equals expected, actual, "The root node has an incorrect name."


    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_correctNameForRoot_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure" & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub


Public Sub Test_vtkExportAsXMLDom_correctNumberOfInfoObject()
'   info -> 1 info object

    On Error GoTo Test_vtkExportAsXMLDom_correctNumberOfInfoObject_Error

    mAssert.Equals countElementsInDom("info", retDom), 1, "DOM should have 1 info object."

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_correctNumberOfInfoObject_Error:
    Err.Source = "Test_vtkExportAsXMLDom_correctNumberOfInfoObject of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub


Public Sub Test_vtkExportAsXMLDom_correctInfoContent()
'   info -> correct content

    On Error GoTo Test_vtkExportAsXMLDom_correctDomContent_Error
    
    mAssert.Equals retDom.getElementsByTagName("info").Item(0).ChildNodes.Item(0).BaseName, "projectName"
    mAssert.Equals retDom.getElementsByTagName("info").Item(0).ChildNodes.Item(0).Text, testProjectName
    
    mAssert.Equals retDom.getElementsByTagName("info").Item(0).ChildNodes.Item(1).BaseName, "vtkConfigurationsVersion"
    mAssert.Equals retDom.getElementsByTagName("info").Item(0).ChildNodes.Item(1).Text, "1.0"

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_correctDomContent_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkExportAsXMLDom_correctDomContent of Class Module vtkXMLExportTester"
    Resume Next
End Sub


Public Sub Test_vtkExportAsXMLDom_correctConfigurationNumber()
'   configurations -> correct number of configurations
    
    On Error GoTo Test_vtkExportAsXMLDom_correctConfigurationNumber_Error
    
    Dim numConf As Integer
    numConf = cm.configurationCount

    mAssert.Equals countElementsInDom("configuration", retDom), numConf, "DOM should have " & numConf & " configuration objects."

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_correctConfigurationNumber_Error:
    Err.Source = "Test_vtkExportAsXMLDom_correctConfigurationNumber of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub


Public Sub Text_vtkExportAsDom_correctConfigurationContent()
'   configurations -> correct content

    On Error GoTo Text_vtkExportAsDom_correctConfigurationContent_Error

    Dim retDom As MSXML2.DOMDocument
    
    Set retDom = vtkExportAsXMLDOM(testWorkbook, testProjectName)
    
    Dim i As Integer
    Dim expected As String
    Dim actual As String
    
    ' For each configuration in the configuration manager
    Dim conf As vtkConfiguration
    For Each conf In cm.configurations
        
        ' Testing the name of the first child node (name)
        expected = "name"
        actual = retDom.getElementsByTagName("configuration").Item(conf.Number - 1).ChildNodes.Item(0).BaseName
        mAssert.Equals expected, actual, "First child node of the configuration in the XML DOM has an incorrect name."

        ' Testing the content of the first child node (name)
        expected = conf.name
        actual = retDom.getElementsByTagName("configuration").Item(conf.Number - 1).ChildNodes.Item(0).Text
        mAssert.Equals expected, actual, "First child node of the configuration in the XML DOM has an incorrect value."
        
        ' Testing the name of the second child node (path)
        expected = "path"
        actual = retDom.getElementsByTagName("configuration").Item(conf.Number - 1).ChildNodes.Item(1).BaseName
        mAssert.Equals expected, actual, "Second child node of the configuration in the XML DOM has an incorrect name."
        
        ' Testing the content of the second child node (path)
        expected = conf.path
        actual = retDom.getElementsByTagName("configuration").Item(conf.Number - 1).ChildNodes.Item(1).Text
        mAssert.Equals expected, actual, "Second child node of the configuration in the XML DOM has an incorrect content."

    Next

    On Error GoTo 0
    Exit Sub

Text_vtkExportAsDom_correctConfigurationContent_Error:
    Err.Source = "Text_vtkExportAsDom_correctConfigurationContent of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub


Public Sub Test_vtkExportAsXMLDom_correctModuleNumber()
'   modules -> correct number of modules

    On Error GoTo Test_vtkExportAsXMLDom_correctModuleNumber_Error
    
    Dim numModules As Integer
    numModules = cm.moduleCount
    
    mAssert.Equals countElementsInDom("configuration", retDom), numModules, "DOM should have " & numModules & " module objects."

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_correctModuleNumber_Error:
    Err.Source = "Test_vtkExportAsXMLDom_correctModuleNumber of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
    Exit Sub
End Sub


Public Sub Test_vtkExportAsDom_correctModuleContent()
'   modules -> correct content
    
    On Error GoTo Text_vtkExportAsDom_correctModuleContent_Error
    
    Dim i As Integer
    Dim expected As String
    Dim actual As String
    
    ' For each module in the configuration manager
    Dim mo As vtkModule
    For Each mo In cm.modules
        
        ' Testing the name of the first child node (name)
        expected = "name"
        actual = retDom.getElementsByTagName("module").Item(mo.Number - 1).ChildNodes.Item(0).BaseName
        mAssert.Equals expected, actual, "First child node of the reference in the XML DOM has an incorrect name."
       
        ' Testing the content of the first child node (name)
        expected = mo.name
        actual = retDom.getElementsByTagName("module").Item(mo.Number - 1).ChildNodes.Item(0).Text
        mAssert.Equals expected, actual, "First child node of the reference in the XML DOM has an incorrect value."
        
        ' Testing the "path" nodes - one for each configuration
        Dim conf As vtkConfiguration
        For Each conf In cm.configurations
        
            ' Testing the name of the node
            expected = "path"
            actual = retDom.getElementsByTagName("module").Item(mo.Number - 1).ChildNodes.Item(conf.Number).BaseName
            mAssert.Equals expected, actual, "Child node #" & conf.Number & " of the reference in the XML DOM has an incorrect name."

            ' Testing the name of the attribute of the node
            expected = "confName"
            actual = retDom.getElementsByTagName("module").Item(mo.Number - 1).ChildNodes.Item(conf.Number).Attributes(0).BaseName
            mAssert.Equals expected, actual, "Child node #" & conf.Number & " of the reference in the XML DOM has an incorrect attribute name."
            
            ' Testing the content of the attribute of the node
            expected = conf.name
            actual = retDom.getElementsByTagName("module").Item(mo.Number - 1).ChildNodes.Item(conf.Number).Attributes(0).Text
            mAssert.Equals expected, actual, "Child node #" & conf.Number & "  the reference in the XML DOM has an incorrect attribute value."

            ' Testing the content of the node
            expected = mo.getPathForConfiguration(conf.name)
            actual = retDom.getElementsByTagName("module").Item(mo.Number - 1).ChildNodes.Item(conf.Number).Text
            mAssert.Equals expected, actual, "Child node #" & conf.Number & "  has an incorrect content."

        Next

    Next

    On Error GoTo 0
    Exit Sub

Text_vtkExportAsDom_correctModuleContent_Error:
    Err.Source = "Text_vtkExportAsDom_correctModuleContent of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    ITestCase_TearDown
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
    Dim fso As New FileSystemObject

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


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctNameForRoot"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctNumberOfInfoObject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctInfoContent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctConfigurationNumber"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctModuleNumber"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsDom_correctModuleContent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_dummyProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkWriteXMLDomToFile_dummyDom"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkWriteXMLDomToFile_fileIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkWriteXMLDomToFile_fileIsOverWritten"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkWriteXMLDomToFile_filepathIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkWriteXMLDomToFile_fileContentIsValidDom"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkExportAsXMLDom_correctNameForRoot": Test_vtkExportAsXMLDom_correctNameForRoot
        Case "Test_vtkExportAsXMLDom_correctNumberOfInfoObject": Test_vtkExportAsXMLDom_correctNumberOfInfoObject
        Case "Test_vtkExportAsXMLDom_correctInfoContent": Test_vtkExportAsXMLDom_correctInfoContent
        Case "Test_vtkExportAsXMLDom_correctConfigurationNumber": Test_vtkExportAsXMLDom_correctConfigurationNumber
        Case "Test_vtkExportAsXMLDom_correctModuleNumber": Test_vtkExportAsXMLDom_correctModuleNumber
        Case "Test_vtkExportAsDom_correctModuleContent": Test_vtkExportAsDom_correctModuleContent
        Case "Test_vtkExportAsXMLDom_dummyProject": Test_vtkExportAsXMLDom_dummyProject
        Case "Test_vtkWriteXMLDomToFile_dummyDom": Test_vtkWriteXMLDomToFile_dummyDom
        Case "Test_vtkWriteXMLDomToFile_fileIsCreated": Test_vtkWriteXMLDomToFile_fileIsCreated
        Case "Test_vtkWriteXMLDomToFile_fileIsOverWritten": Test_vtkWriteXMLDomToFile_fileIsOverWritten
        Case "Test_vtkWriteXMLDomToFile_filepathIsCreated": Test_vtkWriteXMLDomToFile_filepathIsCreated
        Case "Test_vtkWriteXMLDomToFile_fileContentIsValidDom": Test_vtkWriteXMLDomToFile_fileContentIsValidDom
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

