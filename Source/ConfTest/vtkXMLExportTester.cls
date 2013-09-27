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

Private testWorkbook As Workbook
Private cm As vtkConfigurationManager
Private conf As vtkConfiguration


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
'   dummy workbook -> export everything but the references
'   closed workbook -> error VTK_WORKBOOK_NOTOPEN
'   dummy project name -> DOM doesn't have modules or configuration objects
'
' WriteDom
'   dummy DOM
'   file already exists
'   file does not yet exist
'   file content is a valid DOM
'
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


Public Sub Test_vtkExportAsXMLDom_correctNameForRoot()
'   rootNode -> correct name

    On Error GoTo Test_vtkExportAsXMLDom_correctNameForRoot_Error

    Dim retDom As MSXML2.DOMDocument

    Set retDom = vtkExportAsXMLDOM(testWorkbook, testProjectName)

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

    Dim retDom As MSXML2.DOMDocument
    
    Set retDom = vtkExportAsXMLDOM(testWorkbook, testProjectName)

    mAssert.Equals countElementsInDom("info", retDom), 1, "DOM should have 1 info object."

    On Error GoTo 0
    Exit Sub

Test_vtkExportAsXMLDom_correctNumberOfInfoObject_Error:
    Err.source = "Test_vtkExportAsXMLDom_correctNumberOfInfoObject of module vtkXMLExportTester"
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
    ITestCase_TearDown
    Exit Sub
End Sub


'Public Sub Test_vtkExportAsXMLDom_correctDomContent()
''   info -> correct content
'
'    On Error GoTo Test_vtkExportAsXMLDom_correctDomContent_Error
'
'    Dim retDom As MSXML2.DOMDocument
'
'    Set retDom = vtkExportAsXMLDOM(testWorkbook, testProjectName)
'
'    mAssert.Equals retDom.getElementsByTagName("info").Item(1).ChildNodes.Item(1).BaseName, "projectName"
'    mAssert.Equals retDom.getElementsByTagName("info").Item(1).ChildNodes.Item(2).BaseName, "vtkConfigurationsVersion"
'
'    mAssert.Equals retDom.getElementsByTagName("info").Item(1).ChildNodes.Item(1).Text, testProjectName
'    mAssert.Equals retDom.getElementsByTagName("info").Item(1).ChildNodes.Item(2).BaseName, "1.0"
'
'    On Error GoTo 0
'    Exit Sub
'
'Test_vtkExportAsXMLDom_correctDomContent_Error:
'    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkExportAsXMLDom_correctDomContent of Class Module vtkXMLExportTester"
'    Resume Next
'End Sub



Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctNameForRoot"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExportAsXMLDom_correctNumberOfInfoObject"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkExportAsXMLDom_correctNameForRoot": Test_vtkExportAsXMLDom_correctNameForRoot
        Case "Test_vtkExportAsXMLDom_correctNumberOfInfoObject": Test_vtkExportAsXMLDom_correctNumberOfInfoObject
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

