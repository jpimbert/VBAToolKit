VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkXMLUtilitiesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkXMLUtilitiesTester
' Author    : Lucas Vitorino
' Purpose   : Test the utilities contained in the vtkXMLUtilities, except vtkExportAsXML.
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

Private Const listName As String = "projects_list.xml"

Private listFullPath As String
Private fso As New FileSystemObject

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
    listFullPath = fso.BuildPath(VBAToolKit.vtkTestPath, listName)
End Sub

Private Sub ITestCase_TearDown()
    'VBAToolKit.resetTestFolder
End Sub

' Tests to do
'
' vtkCreateXMLListOfRememberedProjects
'   - correct path : file exists
'   - correct path : file content is right
'       - XML version
'       - rootNode name
'   - wrong path : raise error VTK_WRONG_FILE_PATH


Public Sub Test_vtkCreateListOfRememberedProjects_correctPathExistence()
' correct path : file exists

    On Error GoTo Test_vtkCreateListOfRememberedProjects_correctPathExistence_Error
    
    vtkCreateListOfRememberedProjects (listFullPath)
    
    mAssert.Should fso.FileExists(listFullPath), _
                   "File should have been created as the path is correct."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateListOfRememberedProjects_correctPathExistence_Error:
    Err.Source = "Test_vtkCreateXMLListOfRememberedProjects_correctPathExistence of module vtkXMLUtilitiesTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkCreateListOfRememberedProjects_correctPathContent()
' correct path : file content is as expecte

    On Error GoTo Test_vtkCreateListOfRememberedProjects_correctPathContent_Error

    vtkCreateListOfRememberedProjects (listFullPath)
    
    mAssert.Equals vtkTextFileReader(listFullPath), _
                     Chr(10) & _
                     "<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?>" & Chr(10) & _
                     "<rememberedProjects/>", _
                     "List has not been initialized properly."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateListOfRememberedProjects_correctPathContent_Error:
    Err.Source = "Test_vtkCreateListOfRememberedProjects_correctPathContent of module vtkXMLUtilitiesTester"
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub
End Sub



Public Sub Test_vtkCreateListOfRememberedProjects_wrongPathExistence()
' wrong path

    On Error GoTo Test_vtkCreateListOfRememberedProjects_wrongPathExistence_Error

    Dim error As Integer

    listFullPath = fso.BuildPath(VBAToolKit.vtkTestPath, "dummy_folder\list.xml")

    vtkCreateListOfRememberedProjects (listFullPath)

    mAssert.Equals error, VTK_WRONG_FILE_PATH, "An error should have been raised as the folder tree is not valid."

    On Error GoTo 0
    Exit Sub

Test_vtkCreateListOfRememberedProjects_wrongPathExistence_Error:
    error = Err.Number
    Exit Sub
End Sub
Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateListOfRememberedProjects_correctPathExistence"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateListOfRememberedProjects_correctPathContent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkCreateListOfRememberedProjects_wrongPathExistence"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkCreateListOfRememberedProjects_correctPathExistence": Test_vtkCreateListOfRememberedProjects_correctPathExistence
        Case "Test_vtkCreateListOfRememberedProjects_correctPathContent": Test_vtkCreateListOfRememberedProjects_correctPathContent
        Case "Test_vtkCreateListOfRememberedProjects_wrongPathExistence": Test_vtkCreateListOfRememberedProjects_wrongPathExistence
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

