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

    Set retDom = vtkExportAsXMLDOM(testWorkbook, testProjectName)
    
End Sub


Private Sub ITestCase_TearDown()
    On Error Resume Next
    
    vtkResetConfigurationManagers
    vtkCloseAndKillWorkbook wb:=testWorkbook
    vtkCleanFolder VBAToolKit.vtkTestPath
    
End Sub



' -------------------------------------------
' TESTS TO DO
'
' exportDom
'   rootNode -> correct name
'   info -> 1 info object
'   info -> correct content
'   references -> correct number of references
'   references -> correct content
'   conf -> correct number of configurations
'   conf -> correct content
'   modules -> correct number of modules
'   modules -> correct content
'   dummy workbook -> raise VTK_WORKBOOK_NOT_INITIALIZED
'   closed workbook -> raise error VTK_WORKBOOK_NOT_OPEN
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
    Debug.Print "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure countElementsInDom of Class Module vtkXMLExportTester"
    Exit Function
    
End Function

Private Function countElementsInNode(elementName As String, node As MSXML2.IXMLDOMNode) As Integer
    
    Dim count As Integer: count = 0
    
    On Error GoTo countElementsInNode_Error

    Dim subNode As MSXML2.IXMLDOMNode
    For Each subNode In node.ChildNodes
        If StrComp(subNode.BaseName, elementName) = 0 Then count = count + 1
    Next
        
    countElementsInNode = count

    On Error GoTo 0
    Exit Function

countElementsInNode_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure countElementsInNode of Class Module vtkXMLExportTester"
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
    Dim result As String
    expected = "vtkConf"
    result = retDom.ChildNodes.Item(1).BaseName

    mAssert.Should StrComp(expected, result) = 0, _
                    "The root node has an incorrect name." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result   : " & result

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_correctNameForRoot_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure" & Err.source
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
    Err.source = "Test_vtkExportAsXMLDom_correctNumberOfInfoObject of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
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
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportAsXMLDom_correctDomContent of Class Module vtkXMLExportTester"
    Resume Next
End Sub


Public Sub Test_vtkExportAsXMLDom_correctReferenceNumber()
'   references -> correct number of references

    On Error GoTo Test_vtkExportAsXMLDom_correctReferenceNumber_Error
    
    Dim numRef As Integer
    numRef = testWorkbook.VBProject.References.count

    mAssert.Equals countElementsInDom("reference", retDom), numRef, "DOM should have " & numRef & " reference objects."

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_correctReferenceNumber_Error:
    Err.source = "Test_vtkExportAsXMLDom_correctReferenceNumber of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
    ITestCase_TearDown
    Exit Sub
End Sub

Public Sub Test_vtkExportAsXMLDom_correctReferenceContent()
'   references -> correct value of the references object
    
    On Error GoTo Test_vtkExportAsXMLDom_correctReferenceContent_Error
    
    Dim i As Integer
    Dim expected As String
    Dim result As String
    
    For i = 1 To testWorkbook.VBProject.References.count
              
        ' Testing the name of the first child node (name)
        expected = "name"
        result = retDom.getElementsByTagName("reference").Item(i - 1).ChildNodes.Item(0).BaseName
        mAssert.Should StrComp(expected, result) = 0, _
                        "First child node of the reference in the XML DOM has an incorrect name." & vbNewLine & _
                        "Expected   : " & expected & vbNewLine & _
                        "Result     : " & result
        
        ' Testing the content of the first child node (name)
        expected = testWorkbook.VBProject.References(i).name
        result = retDom.getElementsByTagName("reference").Item(i - 1).ChildNodes.Item(0).Text
        mAssert.Should StrComp(expected, result) = 0, _
                        "First child node of the reference in the XML DOM has an incorrect value." & vbNewLine & _
                        "Expected   : " & expected & vbNewLine & _
                        "Result     : " & result
        
        ' Testing the name of the second child node (GUID or path)
        expected = "GUID"
        If testWorkbook.VBProject.References(i).GUID = "" Then expected = "path"
        result = retDom.getElementsByTagName("reference").Item(i - 1).ChildNodes.Item(1).BaseName
        mAssert.Should StrComp(expected, result) = 0, _
                        "Second child node of the reference in the XML DOM has an incorrect name." & vbNewLine & _
                        "Expected   : " & expected & vbNewLine & _
                        "Result     : " & result
        
        ' Testing the content of the second child node (GUID or path)
        expected = testWorkbook.VBProject.References(i).GUID
        If testWorkbook.VBProject.References(i).GUID = "" Then expected = testWorkbook.VBProject.References(i).fullPath
        result = retDom.getElementsByTagName("reference").Item(i - 1).ChildNodes.Item(1).Text
        mAssert.Should StrComp(expected, result) = 0, _
                        "Reference in the XML DOM has an incorrect GUID or Path." & vbNewLine & _
                        "Expected   : " & expected & vbNewLine & _
                        "Result     : " & result

    Next

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_correctReferenceContent_Error:
    Err.source = "Test_vtkExportAsXMLDom_correctReferenceContent of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
    ITestCase_TearDown
    Exit Sub
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
    Err.source = "Test_vtkExportAsXMLDom_correctConfigurationNumber of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
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
    Dim result As String
    
    Dim conf As vtkConfiguration
    For Each conf In cm.configurations
        
        ' Testing the name of the first child node (name)
        expected = "name"
        result = retDom.getElementsByTagName("configuration").Item(i - 1).ChildNodes.Item(0).BaseName
        mAssert.Should StrComp(expected, result) = 0, _
                        "First child node of the reference in the XML DOM has an incorrect name." & vbNewLine & _
                        "Expected   : " & expected & vbNewLine & _
                        "Result     : " & result
        
        ' Testing the content of the first child node (name)
        expected = conf.name
        result = retDom.getElementsByTagName("configuration").Item(i - 1).ChildNodes.Item(0).Text
        mAssert.Should StrComp(expected, result) = 0, _
                        "First child node of the reference in the XML DOM has an incorrect value." & vbNewLine & _
                        "Expected   : " & expected & vbNewLine & _
                        "Result     : " & result
        
        ' Testing the name of the second child node (path)
        expected = "path"
        result = retDom.getElementsByTagName("configuration").Item(i - 1).ChildNodes.Item(1).BaseName
        mAssert.Should StrComp(expected, result) = 0, _
                        "Second child node of the reference in the XML DOM has an incorrect name." & vbNewLine & _
                        "Expected   : " & expected & vbNewLine & _
                        "Result     : " & result
        
        ' Testing the content of the second child node (path)
        expected = conf.path
        If testWorkbook.VBProject.References(i).GUID = "" Then expected = testWorkbook.VBProject.References(i).fullPath
        result = retDom.getElementsByTagName("configuration").Item(i - 1).ChildNodes.Item(1).Text
        mAssert.Should StrComp(expected, result) = 0, _
                        "Reference in the XML DOM has an incorrect GUID or Path." & vbNewLine & _
                        "Expected   : " & expected & vbNewLine & _
                        "Result     : " & result

    Next

    On Error GoTo 0
    Exit Sub

Text_vtkExportAsDom_correctConfigurationContent_Error:
    Err.source = "Text_vtkExportAsDom_correctConfigurationContent of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
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
    Err.source = "Test_vtkExportAsXMLDom_correctModuleNumber of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
    ITestCase_TearDown
    Exit Sub
End Sub


Public Sub Test_vtkExportAsDom_correctModuleContent()
'   modules -> correct content
    
    On Error GoTo Text_vtkExportAsDom_correctModuleContent_Error
    
    Dim i As Integer
    Dim expected As String
    Dim result As String
    
    Dim mo As vtkModule
    For Each mo In cm.modules
        
        ' Testing the name of the first child node (name)
        expected = "name"
        result = retDom.getElementsByTagName("module").Item(mo.number - 1).ChildNodes.Item(0).BaseName

        mAssert.Should StrComp(expected, result) = 0, _
                        "First child node of the reference in the XML DOM has an incorrect name." & vbNewLine & _
                        "Expected   : " & expected & vbNewLine & _
                        "Result     : " & result
        
        ' Testing the content of the first child node (name)
        expected = mo.name
        result = retDom.getElementsByTagName("module").Item(mo.number - 1).ChildNodes.Item(0).Text
        mAssert.Should StrComp(expected, result) = 0, _
                        "First child node of the reference in the XML DOM has an incorrect value." & vbNewLine & _
                        "Expected   : " & expected & vbNewLine & _
                        "Result     : " & result
        
        ' Testing the "path" nodes - one for each configuration
        Dim conf As vtkConfiguration
        For Each conf In cm.configurations
        
            ' Testing the name of the node
            expected = "path"
            result = retDom.getElementsByTagName("module").Item(mo.number - 1).ChildNodes.Item(conf.number).BaseName
            mAssert.Should StrComp(expected, result) = 0, _
                            "Child node #" & conf.number & " of the reference in the XML DOM has an incorrect name." & vbNewLine & _
                            "Expected   : " & expected & vbNewLine & _
                            "Result     : " & result

            ' Testing the name of the attribute of the node
            expected = "confName"
            result = retDom.getElementsByTagName("module").Item(mo.number - 1).ChildNodes.Item(conf.number).Attributes(0).BaseName
            mAssert.Should StrComp(expected, result) = 0, _
                            "Child node #" & conf.number & "  of the reference in the XML DOM has an incorrect attribute name." & vbNewLine & _
                            "Expected   : " & expected & vbNewLine & _
                            "Result     : " & result
            
            ' Testing the content of the attribute of the node
            expected = conf.name
            result = retDom.getElementsByTagName("module").Item(mo.number - 1).ChildNodes.Item(conf.number).Attributes(0).Text
            mAssert.Should StrComp(expected, result) = 0, _
                            "Child node #" & conf.number & "  the reference in the XML DOM has an incorrect attribute value." & vbNewLine & _
                            "Expected   : " & expected & vbNewLine & _
                            "Result     : " & result
                            
            ' Testing the content of the node
            expected = mo.getPathForConfiguration(conf.name)
            result = retDom.getElementsByTagName("module").Item(mo.number - 1).ChildNodes.Item(conf.number).Text
            mAssert.Should StrComp(expected, result) = 0, _
                            "Child node #" & conf.number & "  has an incorrect content." & vbNewLine & _
                            "Expected   : " & expected & vbNewLine & _
                            "Result     : " & result
                            
        Next

    Next

    On Error GoTo 0
    Exit Sub

Text_vtkExportAsDom_correctModuleContent_Error:
    Err.source = "Text_vtkExportAsDom_correctModuleContent of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
    ITestCase_TearDown
    Exit Sub
End Sub

Public Sub Test_vtkExportAsDom_dummyWorkbook()
'   dummy workbook -> raise VTK_WORKBOOK_NOT_INITIALIZED

    On Error GoTo Test_vtkExportAsDom_dummyWorkbook_Error

    Dim dummyWorkbook As Workbook
    Dim error As Integer: error = 0

    Set retDom = vtkExportAsXMLDOM(dummyWorkbook, testProjectName)

    mAssert.Equals error, VTK_WORKBOOK_NOT_INITIALIZED, "Workbook is not initialized : VTK_WORKBOOK_NOT_INITIALIZED should be raised."

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsDom_dummyWorkbook_Error:
    error = Err.number
    Resume Next
End Sub


Public Sub Test_vtkExportAsXMLDom_closedWorkbook()
'   dummy workbook -> raise VTK_WORKBOOK_NOT_OPEN

    On Error GoTo Test_vtkExportAsXMLDom_closedWorkbook_Error

    Dim error As Integer: error = 0
    
    testWorkbook.Close

    Set retDom = vtkExportAsXMLDOM(testWorkbook, testProjectName)

    mAssert.Equals error, VTK_WORKBOOK_NOT_OPEN, "Workbook is not open : VTK_WORKBOOK_NOT_OPEN should be raised."

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_closedWorkbook_Error:
    error = Err.number
    Resume Next
End Sub

Public Sub Test_vtkExportAsXMLDom_dummyProject()
'   dummy project name -> raise VTK_PROJECT_NOT_INITIALIZED

    On Error GoTo Test_vtkExportAsXMLDom_dummyProject_Error
    
    Dim error As Integer: error = 0
    Dim dummyProjectName As String
    dummyProjectName = "dummyProject"
    
    Set retDom = vtkExportAsXMLDOM(testWorkbook, dummyProjectName)
    
    mAssert.Equals error, VTK_PROJECT_NOT_INITIALIZED, "Project is empty : VTK_PROJECT_NOT_INITIALIZED should be raised."

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_dummyProject_Error:
    error = Err.number
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
    
    mAssert.Should (fso.fileExists(testFilePath) = False), "File should not have been created as there has been an error."
    
    On Error GoTo 0
    Exit Sub

Test_vtkWriteXMLDomToFile_dummyDom_Error:
    error = Err.number
    Resume Next
End Sub


Public Sub Test_vtkWriteXMLDomToFile_fileIsCreated()
'   file does not exist -> file is created

    On Error GoTo Test_vtkvtkWriteXMLDomToFile_fileIsCreated_Error

    Dim fso As New FileSystemObject

    vtkWriteXMLDOMToFile retDom, testFilePath
    
    mAssert.Should fso.fileExists(testFilePath), "File should have been created as everything is normal."

    On Error GoTo 0
    Exit Sub

Test_vtkvtkWriteXMLDomToFile_fileIsCreated_Error:
    Err.source = "Test_vtkvtkWriteXMLDomToFile_fileIsCreated of module aavtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
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
    Err.source = "Test_vtkWriteXMLDomToFile_fileIsOverWritten of module aavtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
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
    mAssert.Should Not fso.fileExists(newFilePath), "File should not have been created."

    mAssert.Equals error, VTK_WRONG_FILE_PATH

    On Error GoTo 0
    Exit Sub

Test_vtkWriteXMLDomTOFile_filepathIsCreated_Error:
    error = Err.number
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
    Err.source = "Test_vtkWriteXMLDomToFile_fileContentIsValidDom of module aavtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
    ITestCase_TearDown
    Exit Sub
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctNameForRoot"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctNumberOfInfoObject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctInfoContent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctReferenceNumber"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctReferenceContent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctConfigurationNumber"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctModuleNumber"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsDom_correctModuleContent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsDom_dummyWorkbook"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_closedWorkbook"
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
        Case "Test_vtkExportAsXMLDom_correctReferenceNumber": Test_vtkExportAsXMLDom_correctReferenceNumber
        Case "Test_vtkExportAsXMLDom_correctReferenceContent": Test_vtkExportAsXMLDom_correctReferenceContent
        Case "Test_vtkExportAsXMLDom_correctConfigurationNumber": Test_vtkExportAsXMLDom_correctConfigurationNumber
        Case "Test_vtkExportAsXMLDom_correctModuleNumber": Test_vtkExportAsXMLDom_correctModuleNumber
        Case "Test_vtkExportAsDom_correctModuleContent": Test_vtkExportAsDom_correctModuleContent
        Case "Test_vtkExportAsDom_dummyWorkbook": Test_vtkExportAsDom_dummyWorkbook
        Case "Test_vtkExportAsXMLDom_closedWorkbook": Test_vtkExportAsXMLDom_closedWorkbook
        Case "Test_vtkExportAsXMLDom_dummyProject": Test_vtkExportAsXMLDom_dummyProject
        Case "Test_vtkWriteXMLDomToFile_dummyDom": Test_vtkWriteXMLDomToFile_dummyDom
        Case "Test_vtkWriteXMLDomToFile_fileIsCreated": Test_vtkWriteXMLDomToFile_fileIsCreated
        Case "Test_vtkWriteXMLDomToFile_fileIsOverWritten": Test_vtkWriteXMLDomToFile_fileIsOverWritten
        Case "Test_vtkWriteXMLDomToFile_filepathIsCreated": Test_vtkWriteXMLDomToFile_filepathIsCreated
        Case "Test_vtkWriteXMLDomToFile_fileContentIsValidDom": Test_vtkWriteXMLDomToFile_fileContentIsValidDom
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

