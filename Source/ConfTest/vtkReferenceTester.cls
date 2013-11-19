VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkReferenceTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkReferenceTester
' Author    : Lucas Vitorino
' Purpose   : test the vtkReference object
'
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

Private ref As New vtkReference

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
End Sub

Private Sub ITestCase_TearDown()
    VBAToolKit.resetTestFolder
End Sub

Public Sub TestSetAndGetName()
' test the let and get property of the "name" attribute
    On Error GoTo TestSetAndGetName_Error

    ref.name = "refName"
    mAssert.Equals ref.name, "refName"

    On Error GoTo 0
    Exit Sub

TestSetAndGetName_Error:
    Err.Source = "vtkReferenceTester::TestSetAndGetName"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub TestSetAndGetGUID()
' test the let and get property of the "guid" attribute
    On Error GoTo TestSetAndGetGUID_Error

    ref.guid = "refGUID"
    mAssert.Equals ref.guid, "refGUID"

    On Error GoTo 0
    Exit Sub

TestSetAndGetGUID_Error:
    Err.Source = "vtkReferenceTester::TestSetAndGetGUID"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub TestSetAndGetpath()
' test the let and get property of the "path" attribute
    On Error GoTo TestSetAndGetpath_Error

    ref.path = "refPath"
    mAssert.Equals ref.path, "refPath"

    On Error GoTo 0
    Exit Sub

TestSetAndGetpath_Error:
    Err.Source = "vtkReferenceTester::TestSetAndGetPath"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Public Sub TestGetNameNotInitialized()
' test the return of the get property when it has not been set

    On Error GoTo TestGetNameNotInitialized_Error

    mAssert.Equals ref.name, ""

    On Error GoTo 0
    Exit Sub

TestGetNameNotInitialized_Error:
    Err.Source = "vtkReferenceTester::TestGetNameNotInitialized"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub TestGetGuidNotInitialized()
' test the return of the get property when it has not been set

    On Error GoTo TestGetGuidNotInitialized_Error

    mAssert.Equals ref.guid, ""

    On Error GoTo 0
    Exit Sub

TestGetGuidNotInitialized_Error:
    Err.Source = "vtkReferenceTester::TestGetGuidNotInitialized"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub

Public Sub TestGetpathNotInitialized()
' test the return of the get property when it has not been set

    On Error GoTo TestGetpathNotInitialized_Error

    mAssert.Equals ref.path, ""

    On Error GoTo 0
    Exit Sub

TestGetpathNotInitialized_Error:
    Err.Source = "vtkReferenceTester::TestGetpathNotInitialized"
    mAssert.Should False, "Unexpected Error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
    Exit Sub

End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSetAndGetName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSetAndGetGUID"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSetAndGetpath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetNameNotInitialized"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetGuidNotInitialized"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGetpathNotInitialized"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestSetAndGetName": TestSetAndGetName
        Case "TestSetAndGetGUID": TestSetAndGetGUID
        Case "TestSetAndGetpath": TestSetAndGetpath
        Case "TestGetNameNotInitialized": TestGetNameNotInitialized
        Case "TestGetGuidNotInitialized": TestGetGuidNotInitialized
        Case "TestGetpathNotInitialized": TestGetpathNotInitialized
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

