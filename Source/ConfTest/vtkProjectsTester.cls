VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkProjectsTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkProjectsTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the vtkProjects module
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

Private projectName As String      ' default workbook Name for test

Private fso As New FileSystemObject
Private Const xmlFileName As String = "VBAToolKitTestProjects.xml"
Private Const testProject1Name As String = "TestProject1"
Private Const testProject1XmlRelPath As String = "TestProject1.xml"
Private Const testProject2Name As String = "TestProject2"
Private Const testProject2XmlRelPath As String = "Project\TestProject2.xml"
Private testProject1RootPath As String
Private testProject2RootPath As String

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
    projectName = "TestProject"
    
    ' set the missing consts (calling function so can't be set out of a function/sub)
    testProject1RootPath = fso.BuildPath(VBAToolKit.vtkTestPath, testProject1Name)
    testProject2RootPath = fso.BuildPath(VBAToolKit.vtkTestPath, testProject2Name)
    
    ' Copy a list of projects
    getTestFileFromTemplate fileName:=xmlFileName, openExcel:=False
    
    ' Indicate to the current workbook where is the list of projects
    vtkSetRememberedProjectsFullPath fso.BuildPath(VBAToolKit.vtkTestPath, xmlFileName)
    
End Sub

Private Sub ITestCase_TearDown()
    vtkResetProjects
    resetTestFolder
End Sub

Public Sub TestProjectIsNotNull()
'       Verify that the configuration Manager is not null
    mAssert.Should Not vtkProjectForName(projectName) Is Nothing, "A project must not be null"
End Sub

Public Sub TestProjectsAreDifferentForTwoNames()
'       Verify that the configuration Managers for two project are different
    mAssert.Should Not vtkProjectForName(projectName) Is vtkProjectForName("SecondProject"), "Two projects must not be equal"
End Sub

Public Sub TestProjectIsInvariant()
'       Verify that the configuration Manager for a project is always the same
    Dim cm As vtkProject
    Set cm = vtkProjectForName(projectName)
    mAssert.Should vtkProjectForName(projectName) Is cm, "The project for a project name must be always the same"
End Sub


' Update of vtkProjects
' --------------------
'
' The new functionalities of vtkProjects must be tested
'
' - vtkProjects must read and provide data for a project that is in the file.
' - a sheet with a wrong format is not read and raises an error
' - modifying or deleting the xml file after it has been loaded does not have any effect
' - the export of the sheet must be readable, contain the information and be conform
' - add a project via the "add remembered project" function must also add it to the sheet
' - remove a project must do the same
'

Public Sub Test_vtkProjects_listIsRead()

    On Error GoTo Test_vtkProjects_listIsRead_Error

    mAssert.Equals vtkRootPathForProject("TestProject1"), testProject1RootPath, _
                   "vtkProjects should access to the information in the sheet."
                   
    mAssert.Equals vtkXmlRelPathForProject("TestProject1"), testProject1XmlRelPath, _
                    "vtkProjects should access to the information in the sheet."

    On Error GoTo 0
    Exit Sub

Test_vtkProjects_listIsRead_Error:
    Err.Source = "Sub Test_vtkProjects_listIsRead in module vtkProjectsTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub Test_vtkProjects_listHasWrongFormat()

    On Error GoTo Test_vtkProjects_listHasWrongFormat_Error

    Dim error As Integer: error = 0
    
    

    On Error GoTo 0
    Exit Sub

Test_vtkProjects_listHasWrongFormat_Error:
    Err.Source = "Sub Test_vtkProjects_listHasWrongFormat in module vtkProjectsTester"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectIsNotNull"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectsAreDifferentForTwoNames"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectIsInvariant"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkProjects_listIsRead"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkProjects_listHasWrongFormat"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestProjectIsNotNull": TestProjectIsNotNull
        Case "TestProjectsAreDifferentForTwoNames": TestProjectsAreDifferentForTwoNames
        Case "TestProjectIsInvariant": TestProjectIsInvariant
        Case "Test_vtkProjects_listIsRead": Test_vtkProjects_listIsRead
        Case "Test_vtkProjects_listHasWrongFormat": Test_vtkProjects_listHasWrongFormat
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub



