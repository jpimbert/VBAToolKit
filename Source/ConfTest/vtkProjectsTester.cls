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
End Sub

Private Sub ITestCase_TearDown()
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

'---------------------------------------------------------------------------------------
' Procedure : Test_xmlRememberedProjectsFullPath_GetDefault
' Author    : Champonnois
' Date      : 06/05/2014
' Purpose   : verify the existance file xml VBAToolKitProjects in folder add-in
'---------------------------------------------------------------------------------------
Public Sub Test_xmlRememberedProjectsFullPath_GetDefault()
    Dim fso As New FileSystemObject
    
    On Error GoTo Test_xmlRememberedProjectsFullPath_GetDefault_Error
    
    mAssert.Equals fso.FileExists(xmlRememberedProjectsFullPath), True, "Xml file recover default folder add-ins not present"

    On Error GoTo 0

    Exit Sub

Test_xmlRememberedProjectsFullPath_GetDefault_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_xmlRememberedProjectsFullPath_GetDefault_Error of Class Module vtkProjectsTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_xmlRememberedProjectsFullPath_GetDefaultFileXmlNonexistent
' Author    : Champonnois
' Date      : 06/05/2014
' Purpose   : verify the nonexistence file xml VBAToolKitProjects in folder add-in
'---------------------------------------------------------------------------------------
Public Sub Test_xmlRememberedProjectsFullPath_GetDefaultFileXmlNonexistent()
    Dim fso As New FileSystemObject
    
    On Error GoTo Test_xmlRememberedProjectsFullPath_GetDefaultFileXmlNonexistent_Error
    
    fso.MoveFile xmlRememberedProjectsFullPath, vtkTestPath & "\" & "VBAToolKitProjects.xml"
    
    mAssert.Equals fso.FileExists(xmlRememberedProjectsFullPath), False, "Xml file recover default add-in is present"

    fso.MoveFile vtkTestPath & "\" & "VBAToolKitProjects.xml", xmlRememberedProjectsFullPath

    On Error GoTo 0

    Exit Sub

Test_xmlRememberedProjectsFullPath_GetDefaultFileXmlNonexistent_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_xmlRememberedProjectsFullPath_GetDefaultFileXmlNonexistent of Class Module vtkProjectsTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_isXmlSheetValid_FileXmlExistent
' Author    : Champonnois
' Date      : 06/05/2014
' Purpose   : verify the load file xml VBAToolKitProjects in folder add-in
'---------------------------------------------------------------------------------------
Public Sub Test_isXmlSheetValid_FileXmlExistent()
    Dim fso As New FileSystemObject
    
    On Error GoTo Test_isXmlSheetValid_FileXmlExistent_Error
    
    mAssert.Equals isXmlSheetValid, True, "Xml file not valid or nonexistent"

    On Error GoTo 0

    Exit Sub

Test_isXmlSheetValid_FileXmlExistent_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_isXmlSheetValid_FileXmlExistent of Class Module vtkProjectsTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_isXmlSheetValid_FileXmlNonexistent
' Author    : Champonnois
' Date      : 06/05/2014
' Purpose   : verify the nonexistence load file xml VBAToolKitProjects in folder add-in
'---------------------------------------------------------------------------------------
Public Sub Test_isXmlSheetValid_FileXmlNonexistent()
    Dim fso As New FileSystemObject
    
    On Error GoTo Test_isXmlSheetValid_FileXmlNonexistent_Error
    
    fso.MoveFile xmlRememberedProjectsFullPath, vtkTestPath & "\" & "VBAToolKitProjects.xml"
    
    mAssert.Equals isXmlSheetValid, False, "Xml file is valid and present"
    
    fso.MoveFile vtkTestPath & "\" & "VBAToolKitProjects.xml", xmlRememberedProjectsFullPath

    On Error GoTo 0

    Exit Sub

Test_isXmlSheetValid_FileXmlNonexistent_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_isXmlSheetValid_FileXmlNonexistent of Class Module vtkProjectsTester"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Test_isXmlSheetValid_FileXmlInvalid
' Author    : Champonnois
' Date      : 06/05/2014
' Purpose   : verify the invalid load file txt VBAToolKitProjects in folder add-in
'---------------------------------------------------------------------------------------
Public Sub Test_isXmlSheetValid_FileXmlInvalid()
    Dim fso As New FileSystemObject
    
    On Error GoTo Test_isXmlSheetValid_FileXmlInvalid_Error
    
    fso.MoveFile xmlRememberedProjectsFullPath, vtkTestPath & "\" & "VBAToolKitProjects.xml"
    fso.CreateTextFile (fso.GetParentFolderName(xmlRememberedProjectsFullPath) & "\" & "VBAToolKitProjects.pdf")
    
    mAssert.Equals isXmlSheetValid, False, "Xml file is valid, but it shouldn't"
    
    fso.DeleteFile (fso.GetParentFolderName(xmlRememberedProjectsFullPath) & "\" & "VBAToolKitProjects.pdf")
    fso.MoveFile vtkTestPath & "\" & "VBAToolKitProjects.xml", xmlRememberedProjectsFullPath

    On Error GoTo 0

    Exit Sub

Test_isXmlSheetValid_FileXmlInvalid_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_isXmlSheetValid_FileXmlInvalid of Class Module vtkProjectsTester"
    Resume Next
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectIsNotNull"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectsAreDifferentForTwoNames"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectIsInvariant"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_xmlRememberedProjectsFullPath_GetDefault"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_xmlRememberedProjectsFullPath_GetDefaultFileXmlNonexistent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_isXmlSheetValid_FileXmlExistent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_isXmlSheetValid_FileXmlNonexistent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_isXmlSheetValid_FileXmlInvalid"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestProjectIsNotNull": TestProjectIsNotNull
        Case "TestProjectsAreDifferentForTwoNames": TestProjectsAreDifferentForTwoNames
        Case "TestProjectIsInvariant": TestProjectIsInvariant
        Case "Test_xmlRememberedProjectsFullPath_GetDefault": Test_xmlRememberedProjectsFullPath_GetDefault
        Case "Test_xmlRememberedProjectsFullPath_GetDefaultFileXmlNonexistent": Test_xmlRememberedProjectsFullPath_GetDefaultFileXmlNonexistent
        Case "Test_isXmlSheetValid_FileXmlExistent": Test_isXmlSheetValid_FileXmlExistent
        Case "Test_isXmlSheetValid_FileXmlNonexistent": Test_isXmlSheetValid_FileXmlNonexistent
        Case "Test_isXmlSheetValid_FileXmlInvalid": Test_isXmlSheetValid_FileXmlInvalid
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub



