VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkNormalizeTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkNormalizeTester
' Author    : Lucas Vitorino
' Purpose   : Test the functions in the vtkNormalize module.
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

Private listOfIdentifiersToNormalize As String

Private Sub initializeListOfIdentifiersToNormalize()
    listOfWordsToNormalize = _
    "FooBar" & "," & _
    "test1" & "," & _
    "TEST2" & "," & _
    "foo_BAR" & "," & _
    "foo-bar" & "," & _
    "1Test" & "," & _
    "Foo|Bar" & "," & _
    ""
    'The 3 last ones are not proper identifiers : no hyphen, no starting by a number, no special charachters
End Sub

Private Function getListOfIdentifiersToNormalize() As String()
    initializeListOfIdentifiersToNormalize
    getListOfWordsToNormalize = Split(listOfWordsToNormalize, ",")
End Function


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

End Sub

' TESTS TO DO
' Normalize token
'   token in list, wrong case : case is proper
'   token in list, wrong case : case is proper
'   token containing a token in list : case unchanged
'   token not in list : not modified
'   list is empty : no token modified
'
' Normalize String
'   tokens in list, no comment : case corrected
'   tokens in list after comment : case is corrected before comment but not after
'   tokens are not identifiers : their case should not be changed
'   test special characters like vbTab and vbNewLine : should not be different
'
' Normalize File
'   wrong file path : VTK_WRONG_PATH
'   file has the same length before and after normalization
'   file content is properly cased


Public Sub Test_vtkNormalizeToken_wrongCase1()
    
    Dim normalized As String
    
    On Error GoTo Test_vtkNormalizeToken_wrongCase1_Error

    mAssert.Equals StrComp(vtkNormalizeToken("FOOBAR", getListOfWordsToNormalize), "FooBar"), 0

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_wrongCase1_Error:
    Debug.Print "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_wrongCase1 of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeToken_wrongCase2()

    On Error GoTo Test_vtkNormalizeToken_wrongCase2_Error

    mAssert.Equals StrComp(vtkNormalizeToken("teSt2", getListOfWordsToNormalize), "TEST2"), 0

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_wrongCase2_Error:
    Debug.Print "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_wrongCase2 of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeToken_wrongCase3()

End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeToken_wrongCase1"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkNormalizeToken_wrongCase1": Test_vtkNormalizeToken_wrongCase1
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

