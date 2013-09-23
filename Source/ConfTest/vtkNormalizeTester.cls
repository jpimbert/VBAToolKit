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

Private Const fileLine1 As String = "FOOBAR foobar"
Private Const fileLine2 As String = "FOO-BAR TEST1"
Private dummyFile As String

Private listOfIdentifiersToNormalize As String

Private Sub initializeListOfIdentifiersToNormalize()
    listOfIdentifiersToNormalize = _
    "FooBar" & "," & _
    "test1" & "," & _
    "TEST2" & "," & _
    "foo_BAR" & "," & _
    "foo-bar" & "," & _
    "1Test" & "," & _
    ""
    'The 3 last ones are not proper identifiers : no hyphen, no starting by a number, no special characters
End Sub

Private Function getListOfIdentifiersToNormalize() As String()
    initializeListOfIdentifiersToNormalize
    getListOfIdentifiersToNormalize = Split(listOfIdentifiersToNormalize, ",")
End Function

Private Function fillFile(filePath As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream

    ts = fso.OpenTextFile(filePath, ForWriting)
    
    ts.WriteLine fileLine1
    ts.WriteLine fileLine2

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
    
    dummyFile = VBAToolKit.vtkTestPath & "\" & "dummyFile.txt"
    
End Sub

Private Sub ITestCase_TearDown()

    vtkCleanFolder VBAToolKit.vtkTestPath

End Sub

' TESTS TO DO
' Normalize token
'   token in list, wrong case : casing is made proper
'   token in list, wrong case : casing is made proper
'   token containing a token in the list : casing is unchanged
'   token not in list : not modified
'   token not an identifier, but in list : casing is changed
'   list is empty : no token modified
'
' Normalize String
'   tokens in list, no comment : casing is corrected
'   tokens in list after comment : casing is corrected before comment but not after
'   tokens in list between quotes and after : casing is corrected before the comment quotes but not after they are closed
'   tokens are not identifiers : their casing should not be changed
'   test special characters like vbTab and vbNewLine : should not be different
'
' Normalize File
'   wrong file path : VTK_WRONG_FILE_PATH
'   file has the same length before and after normalization
'   file content is properly cased
'

Public Sub Test_vtkNormalizeToken_wrongCase1()
'   token in list, wrong case : casing is made proper
    
    Dim normalized As String
    
    On Error GoTo Test_vtkNormalizeToken_wrongCase1_Error

    mAssert.Equals StrComp(vtkNormalizeToken("FOOBAR", getListOfIdentifiersToNormalize), "FooBar"), 0, "Casing should have been corrected as token is in the list."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_wrongCase1_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_wrongCase1 of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeToken_wrongCase2()
'   token in list, wrong case : casing is made proper

    On Error GoTo Test_vtkNormalizeToken_wrongCase2_Error

    mAssert.Equals StrComp(vtkNormalizeToken("teSt2", getListOfIdentifiersToNormalize), "TEST2"), 0, "Casing should have been corrected as token is in the list."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_wrongCase2_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_wrongCase2 of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeToken_wrongCase3()
'   token containing a token in the list : casing is unchanged

    On Error GoTo Test_vtkNormalizeToken_wrongCase3_Error

    mAssert.Equals StrComp(vtkNormalizeToken("FOOBAR foobar", getListOfIdentifiersToNormalize), "FooBar_foobar"), 0, "Casing should not have been corrected as the two tokens are not separated."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_wrongCase3_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_wrongCase3 of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeToken_tokenNotInList()
'   token not in list : not modified

    On Error GoTo Test_vtkNormalizeToken_tokenNotInList_Error

    mAssert.Equals StrComp(vtkNormalizeToken("NotInList", getListOfIdentifiersToNormalize), "NotInList"), 0, "Casing should not have been corrected as token is not in the list."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_tokenNotInList_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_tokenNotInList of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeToken_notAnIdentifierButInList()
'   token not an identifier, but in list : casing is changed

    On Error GoTo Test_vtkNormalizeToken_notAnIdentifierButInList_Error

    mAssert.Equals StrComp(vtkNormalizeToken("1TEST", getListOfIdentifiersToNormalize), "1Test"), 0, "Casing should have been corrected as token has not been filtered yet."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_notAnIdentifierButInList_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_notAnIdentifierButInList of Class Module vtkNormalizeTester"
    Resume Next
End Sub


Public Sub Test_vtkNormalizeToken_listIsEmpty()
'   list is empty : no token modified

    Dim emptyList() As String

    On Error GoTo Test_vtkNormalizeToken_listIsEmpty_Error

    mAssert.Equals StrComp(vtkNormalizeToken("FOOBAR", emptyList), "FOOBAR"), 0, "Casing should be different as list is empty."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_listIsEmpty_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_listIsEmpty of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeString_tokensInListAreCorrected()
'   tokens in list, no comment : casing is corrected

    On Error GoTo Test_vtkNormalizeString_tokensInListAreCorrected_Error

    mAssert.Equals StrComp(vtkNormalizeString("FOOBAR foo_bar", getListOfIdentifiersToNormalize), "FooBar foo_BAR"), 0

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_tokensInListAreCorrected_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_tokensInListAreCorrected of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected()
'   tokens in list after comment : casing is corrected before comment but not after

    On Error GoTo Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected_Error

    mAssert.Equals StrComp(vtkNormalizeString("FOOBAR foo_bar 'FOOBAR", getListOfIdentifiersToNormalize), "FooBar foo_BAR 'FOOBAR"), 0, "Token after comment should not be corrected."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected of Class Module vtkNormalizeTester"
    Resume Next
End Sub


Public Sub Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments()
'   tokens in list between quotes and after : casing is corrected before the comment quotes but not after they are closed

    On Error GoTo Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments_Error

    mAssert.Equals StrComp(vtkNormalizeString("FOOBAR foo_bar 'FOOBAR' FOOBAR", getListOfIdentifiersToNormalize), "FooBar foo_BAR 'FOOBAR' FOOBAR"), 0, "Code after the closing quotes is a comment too and its case should not be changed."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeString_tokensAreNotIdentifiers()
'   tokens are not identifiers : their casing should not be changed

    On Error GoTo Test_vtkNormalizeString_tokensAreNotIdentifiers_Error

    mAssert.Equals StrComp(vtkNormalizeString("FOOBAR fOO-bar 1TEST", getListOfIdentifiersToNormalize), "FooBar fOO-bar 1TEST"), 0, "Only proper identifiers should have been corrected."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_tokensAreNotIdentifiers_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_tokensAreNotIdentifiers of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeString_stringHasSpecialCharacters()
'   test special characters like vbTab and vbNewLine : should not be different
    On Error GoTo Test_vtkNormalizeString_stringHasSpecialCharacters_Error

    mAssert.Equals StrComp( _
    vtkNormalizeString("FOOBAR" & vbNewLine & "TEST1" & vbTab & "foo_bar", getListOfIdentifiersToNormalize), _
    "FooBar" & vbNewLine & "Test1" & vbTab & "foo_BAR"), 0, _
    "No problem should occur with these special characters."
    
    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_stringHasSpecialCharacters_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_stringHasSpecialCharacters of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeFile_wrongPath()
'   wrong file path : VTK_WRONG_FILE_PATH
    
    Dim error As Integer: error = 0

    On Error GoTo Test_vtkNormalizeFile_wrongPath_Error

    vtkNormalizeFile "dummyPath", getListOfIdentifiersToNormalize
    
    mAssert.Equals error, VTK_WRONG_FILE_PATH

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeFile_wrongPath_Error:
    error = Err.number
    Resume Next
End Sub

Public Sub Test_vtkNormalizeFile_sameSizeBeforeThanAfter()
'   file has the same length before and after normalization

    On Error GoTo Test_vtkNormalizeFile_sameSizeBeforeThanAfter_Error

    Dim fso As New FileSystemObject

    Dim oldSize As Long: oldSize = 0
    Dim newSize As Long: newSize = 0
    
    ' Creation and initialization of the file
    fso.CreateTextFile dummyFile
    fillFile dummyFile

    ' Measure size before processing, process, measure size after processing
    oldSize = fso.GetFile(dummyFile).Size
    vtkNormalizeFile dummyFile, getListOfIdentifiersToNormalize
    newSize = fso.GetFile(dummyFile).Size
    
    ' Compare
    mAssert.Equals oldSize, newSize, "Normalization shouldn't change the file size."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeFile_sameSizeBeforeThanAfter_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeFile_sameSizeBeforeThanAfter of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeFile_fileContent()
'   file content is properly cased

    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    On Error GoTo Test_vtkNormalizeFile_fileContent_Error

    fso.CreateTextFile dummyFile
    fillFile dummyFile
    
    mAssert.Equals StrComp(vtkNormalizeString(fileLine1, getListOfIdentifiersToNormalize), _
                    vtkNormalizeString(ts.ReadLine, getListOfIdentifiersToNormalize)), _
                    0, "First line of the file should be properly cased."

    mAssert.Equals StrComp(vtkNormalizeString(fileLine2, getListOfIdentifiersToNormalize), _
                    vtkNormalizeString(ts.ReadLine, getListOfIdentifiersToNormalize)), _
                    0, "Second line of the file should be properly cased."

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeFile_fileContent_Error:
    mAssert.Should False, "Unexpected error " & Err.number & " (" & Err.Description & ") in procedure Test_vtkNormalizeFile_fileContent of Class Module vtkNormalizeTester"
    Resume Next

End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeToken_wrongCase1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeToken_wrongCase2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeToken_wrongCase3"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeToken_tokenNotInList"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeToken_notAnIdentifierButInList"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeToken_listIsEmpty"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeString_tokensInListAreCorrected"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeString_tokensAreNotIdentifiers"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeString_stringHasSpecialCharacters"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeFile_wrongPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeFile_sameSizeBeforeThanAfter"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeFile_fileContent"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkNormalizeToken_wrongCase1": Test_vtkNormalizeToken_wrongCase1
        Case "Test_vtkNormalizeToken_wrongCase2": Test_vtkNormalizeToken_wrongCase2
        Case "Test_vtkNormalizeToken_wrongCase3": Test_vtkNormalizeToken_wrongCase3
        Case "Test_vtkNormalizeToken_tokenNotInList": Test_vtkNormalizeToken_tokenNotInList
        Case "Test_vtkNormalizeToken_notAnIdentifierButInList": Test_vtkNormalizeToken_notAnIdentifierButInList
        Case "Test_vtkNormalizeToken_listIsEmpty": Test_vtkNormalizeToken_listIsEmpty
        Case "Test_vtkNormalizeString_tokensInListAreCorrected": Test_vtkNormalizeString_tokensInListAreCorrected
        Case "Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected": Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected
        Case "Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments": Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments
        Case "Test_vtkNormalizeString_tokensAreNotIdentifiers": Test_vtkNormalizeString_tokensAreNotIdentifiers
        Case "Test_vtkNormalizeString_stringHasSpecialCharacters": Test_vtkNormalizeString_stringHasSpecialCharacters
        Case "Test_vtkNormalizeFile_wrongPath": Test_vtkNormalizeFile_wrongPath
        Case "Test_vtkNormalizeFile_sameSizeBeforeThanAfter": Test_vtkNormalizeFile_sameSizeBeforeThanAfter
        Case "Test_vtkNormalizeFile_fileContent": Test_vtkNormalizeFile_fileContent
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

