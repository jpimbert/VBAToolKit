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

Private Function fillFile1(filePath As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = fso.OpenTextFile(filePath, ForWriting)
    
    ts.WriteLine fileLine1
    ts.WriteLine fileLine2
    
    ts.Close

End Function

Private Function fillFile2(filePath As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = fso.OpenTextFile(filePath, ForWriting)
    
    ts.Write fileLine1 & vbNewLine
    ts.Write fileLine2
    
    ts.Close

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
'   tokens in list between and out quotes : those between quotes are not modified, those out of them are
'
' Normalize File
'   file doesn't exist : VTK_WRONG_FILE_PATH
'   file is a folder : VTK_FORBIDDEN_PARAMETER
'   file has been initialized with WriteLine : same length before and after normalization
'   file does not end with vbNewLine : after normalization, file has the same length + 2 has a vbNewLine has been added
'   file content is properly cased
'

Public Sub Test_vtkNormalizeToken_wrongCase1()
'   token in list, wrong case : casing is made proper
    
    Dim str As String
    Dim expected As String
    Dim result As String
    
    On Error GoTo Test_vtkNormalizeToken_wrongCase1_Error
    
    str = "FOOBAR"
    expected = "FooBar"
    result = vtkNormalizeToken(str, getListOfIdentifiersToNormalize)
    
    mAssert.Should StrComp(expected, result) = 0, _
                    "Casing should have been corrected as token is in the list." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_wrongCase1_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_wrongCase1 of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeToken_wrongCase2()
'   token in list, wrong case : casing is made proper

    Dim str As String
    Dim expected As String
    Dim result As String

    On Error GoTo Test_vtkNormalizeToken_wrongCase2_Error
    
    str = "teSt2"
    expected = "TEST2"
    result = vtkNormalizeToken(str, getListOfIdentifiersToNormalize)

    mAssert.Should StrComp(expected, result) = 0, _
                    "Casing should have been corrected as token is in the list." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result
                    
    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_wrongCase2_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_wrongCase2 of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeToken_wrongCase3()
'   token containing a token in the list : casing is unchanged
    
    Dim str As String
    Dim expected As String
    Dim result As String

    On Error GoTo Test_vtkNormalizeToken_wrongCase3_Error

    str = "FOOBAR_foobar"
    expected = "FOOBAR_foobar"
    result = vtkNormalizeToken(str, getListOfIdentifiersToNormalize)
    
    mAssert.Should StrComp(expected, result) = 0, _
                    "Casing should have been corrected as token is in the list." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result
                    
    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_wrongCase3_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_wrongCase3 of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeToken_tokenNotInList()
'   token not in list : not modified

    Dim str As String
    Dim expected As String
    Dim result As String

    On Error GoTo Test_vtkNormalizeToken_tokenNotInList_Error

    str = "NotInList"
    expected = "NotInList"
    result = vtkNormalizeToken(str, getListOfIdentifiersToNormalize)

    mAssert.Should StrComp(expected, result) = 0, _
                    "Casing should not have been corrected as token is not in the list." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_tokenNotInList_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_tokenNotInList of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeToken_notAnIdentifierButInList()
'   token not an identifier, but in list : casing is changed

    Dim str As String
    Dim expected As String
    Dim result As String

    On Error GoTo Test_vtkNormalizeToken_notAnIdentifierButInList_Error

    str = "1TEST"
    expected = "1Test"
    result = vtkNormalizeToken(str, getListOfIdentifiersToNormalize)

    mAssert.Should StrComp(expected, result) = 0, _
                    "Casing should have been corrected as token is in the list." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result
    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_notAnIdentifierButInList_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_notAnIdentifierButInList of Class Module vtkNormalizeTester"
    Resume Next
End Sub


Public Sub Test_vtkNormalizeToken_listIsEmpty()
'   list is empty : no token modified

    Dim str As String
    Dim expected As String
    Dim result As String

    Dim emptyList() As String

    On Error GoTo Test_vtkNormalizeToken_listIsEmpty_Error
    
    str = "FOOBAR"
    expected = "FOOBAR"
    result = vtkNormalizeToken(expected, emptyList)

    mAssert.Should StrComp(expected, result) = 0, _
                    "Casing should not be different as list is empty." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result
    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeToken_listIsEmpty_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeToken_listIsEmpty of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeString_tokensInListAreCorrected()
'   tokens in list, no comment : casing is corrected

    Dim str As String
    Dim expected As String
    Dim result As String

    On Error GoTo Test_vtkNormalizeString_tokensInListAreCorrected_Error

    str = "FOOBAR foo_bar"
    expected = "FooBar foo_BAR"
    result = vtkNormalizeString(str, getListOfIdentifiersToNormalize)

    mAssert.Should StrComp(expected, result) = 0, _
                    "Casing should have been corrected as tokens are in the list." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_tokensInListAreCorrected_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_tokensInListAreCorrected of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected()
'   tokens in list after comment : casing is corrected before comment but not after
    
    Dim str As String
    Dim expected As String
    Dim result As String

    On Error GoTo Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected_Error

    str = "FOOBAR foo_bar 'FOOBAR"
    expected = "FooBar foo_BAR 'FOOBAR"
    result = vtkNormalizeString(str, getListOfIdentifiersToNormalize)

    mAssert.Should StrComp(expected, result) = 0, _
                    "Token after comment should not be corrected." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_tokensAfterCommentAreNotCorrected of Class Module vtkNormalizeTester"
    Resume Next
End Sub


Public Sub Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments()
'   tokens in list between quotes and after : casing is corrected before the comment quotes but not after they are closed

    Dim str As String
    Dim expected As String
    Dim result As String

    On Error GoTo Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments_Error

    str = "FOOBAR foo_bar 'FOOBAR' FOOBAR"
    expected = "FooBar foo_BAR 'FOOBAR' FOOBAR"
    result = vtkNormalizeString(str, getListOfIdentifiersToNormalize)

    mAssert.Should StrComp(expected, result) = 0, _
                    "Code after the closing quotes is a comment too and its case should not be changed." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_noBetweenQuotesOnlyAfterComments of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeString_tokensAreNotIdentifiers()
'   tokens are not identifiers : their casing should not be changed

    Dim str As String
    Dim expected As String
    Dim result As String

    On Error GoTo Test_vtkNormalizeString_tokensAreNotIdentifiers_Error

    str = "FOOBAR fOO-bar 1TEST"
    expected = "FooBar fOO-bar 1TEST"
    result = vtkNormalizeString(str, getListOfIdentifiersToNormalize)

    mAssert.Should StrComp(expected, result) = 0, _
                    "Only proper identifiers should have been corrected." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_tokensAreNotIdentifiers_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_tokensAreNotIdentifiers of Class Module vtkNormalizeTester"
    Resume Next
End Sub

Public Sub Test_vtkNormalizeString_stringHasSpecialCharacters()
'   test special characters like vbTab and vbNewLine : should not be different
    
    Dim str As String
    Dim expected As String
    Dim result As String
    
    On Error GoTo Test_vtkNormalizeString_stringHasSpecialCharacters_Error

    str = "FOOBAR" & vbNewLine & "TEST1" & vbTab & "foo_bar"
    expected = "FooBar" & vbNewLine & "test1" & vbTab & "foo_BAR"
    result = vtkNormalizeString(str, getListOfIdentifiersToNormalize)

    mAssert.Should StrComp(expected, result) = 0, _
                    "No problem should occur with these special characters." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_stringHasSpecialCharacters_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_stringHasSpecialCharacters of Class Module vtkNormalizeTester"
    Resume Next
End Sub


Public Sub Test_vtkNormalizeString_quotes()
'   tokens in list between and out quotes : those between quotes are not modified, those out of them are

    On Error GoTo Test_vtkNormalizeString_quotes_Error

    Dim str As String
    Dim expected As String
    Dim result As String
    
    str = "FOOBAR ""foobar"" FOOBAR"
    expected = "FooBar ""foobar"" FooBar"
    result = vtkNormalizeString(str, getListOfIdentifiersToNormalize)

    mAssert.Should StrComp(expected, result) = 0, _
                    "The whole string except the part between the quotes should be normalized." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeString_quotes_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeString_quotes of Class Module vtkNormalizeTester"
    Resume Next
End Sub


Public Sub Test_vtkNormalizeFile_wrongPath()
'   file doesn't exist : VTK_WRONG_FILE_PATH
    
    Dim error As Integer: error = 0

    On Error GoTo Test_vtkNormalizeFile_wrongPath_Error

    vtkNormalizeFile "dummyPath", getListOfIdentifiersToNormalize
    
    mAssert.Equals error, VTK_WRONG_FILE_PATH

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeFile_wrongPath_Error:
    error = Err.Number
    Resume Next
End Sub

Public Sub Test_vtkNormalizeFile_folderPath()
'   file is a folder : VTK_WRONG_FILE_PATH
    
    Dim error As Integer: error = 0
    Dim fso As New FileSystemObject
    Dim dummyFolder As String
    dummyFolder = VBAToolKit.vtkTestPath & "\dummyFolder"
    fso.CreateFolder (dummyFolder)
    
    On Error GoTo Test_vtkNormalizeFile_wrongPath_Error

    vtkNormalizeFile dummyFolder, getListOfIdentifiersToNormalize
    
    mAssert.Equals error, VTK_WRONG_FILE_PATH

    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeFile_wrongPath_Error:
    error = Err.Number
    Resume Next
End Sub

Public Sub Test_vtkNormalizeFile_sameSizeBeforeThanAfter1()
'   file has been initialized with WriteLine : same length before and after normalization

    On Error GoTo Test_vtkNormalizeFile_sameSizeBeforeThanAfter_Error

    Dim fso As New FileSystemObject
    
    Dim oldSize As Long: oldSize = 0
    Dim newSize As Long: newSize = 0
    
    ' Creation and initialization of the file
    fso.CreateTextFile dummyFile
    fillFile1 dummyFile

    ' Measure size before processing, process, measure size after processing
    oldSize = fso.GetFile(dummyFile).Size
    vtkNormalizeFile dummyFile, getListOfIdentifiersToNormalize
    newSize = fso.GetFile(dummyFile).Size
    
    ' Compare
    mAssert.Equals oldSize, newSize, "Normalization shouldn't change the file size as it has a carriage return as las character."
    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeFile_sameSizeBeforeThanAfter_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeFile_sameSizeBeforeThanAfter of Class Module vtkNormalizeTester"
    Exit Sub
End Sub

Public Sub Test_vtkNormalizeFile_sameSizeBeforeThanAfter2()
'   file does not end with vbNewLine : after normalization, file has the same length + 2 has a vbNewLine has been added

    On Error GoTo Test_vtkNormalizeFile_sameSizeBeforeThanAfter_Error

    Dim fso As New FileSystemObject
    
    Dim oldSize As Long: oldSize = 0
    Dim newSize As Long: newSize = 0
    
    ' Creation and initialization of the file
    fso.CreateTextFile dummyFile
    fillFile2 dummyFile

    ' Measure size before processing, process, measure size after processing
    oldSize = fso.GetFile(dummyFile).Size
    vtkNormalizeFile dummyFile, getListOfIdentifiersToNormalize
    newSize = fso.GetFile(dummyFile).Size
    
    ' Compare
    mAssert.Equals oldSize, newSize - 2, "Normalization should add 2 bytes as the file did not have a carriage return as last character."
    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeFile_sameSizeBeforeThanAfter_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeFile_sameSizeBeforeThanAfter of Class Module vtkNormalizeTester"
    Exit Sub
End Sub

Public Sub Test_vtkNormalizeFile_fileContent()
'   file content is properly cased

    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    Dim expected As String
    Dim result As String
    
    On Error GoTo Test_vtkNormalizeFile_fileContent_Error

    ' Initialize the file
    fso.CreateTextFile dummyFile
    fillFile1 dummyFile
    
    ' Process it
    vtkNormalizeFile dummyFile, getListOfIdentifiersToNormalize

    Set ts = fso.OpenTextFile(dummyFile, ForReading)
    
    ' First line
    expected = vtkNormalizeString(fileLine1, getListOfIdentifiersToNormalize)
    result = ts.ReadLine
    mAssert.Should StrComp(expected, result) = 0, _
                    "First line of the file should be properly cased." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result

    ' Second line
    expected = vtkNormalizeString(fileLine2, getListOfIdentifiersToNormalize)
    result = ts.ReadLine
    mAssert.Should StrComp(expected, result) = 0, _
                    "First line of the file should be properly cased." & vbNewLine & _
                    "Expected : " & expected & vbNewLine & _
                    "Result : " & result
    
    ts.Close
        
    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeFile_fileContent_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeFile_fileContent of Class Module vtkNormalizeTester"
    Resume Next

End Sub

Public Sub Test_vtkNormalizeFile_onlyOneFileInFolder()
' tmp file has been deleted

    Dim fso As New FileSystemObject
    
    ' Creation and initialization of the file
    fso.CreateTextFile dummyFile
    fillFile1 dummyFile
    
    On Error GoTo Test_vtkNormalizeFile_onlyOneFileInFolder_Error

    vtkNormalizeFile dummyFile, getListOfIdentifiersToNormalize

    mAssert.Equals fso.GetFolder(VBAToolKit.vtkTestPath).Files.Count, 1, _
            "There must only be one file in the directory, no tmp file or old file."
    
    On Error GoTo 0
    Exit Sub

Test_vtkNormalizeFile_onlyOneFileInFolder_Error:
    mAssert.Should False, "Unexpected error " & Err.Number & " (" & Err.Description & ") in procedure Test_vtkNormalizeFile_onlyOneFileInFolder of Class Module vtkNormalizeTester"
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
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeString_quotes"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeFile_wrongPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeFile_folderPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeFile_sameSizeBeforeThanAfter1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeFile_sameSizeBeforeThanAfter2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeFile_fileContent"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkNormalizeFile_onlyOneFileInFolder"
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
        Case "Test_vtkNormalizeString_quotes": Test_vtkNormalizeString_quotes
        Case "Test_vtkNormalizeFile_wrongPath": Test_vtkNormalizeFile_wrongPath
        Case "Test_vtkNormalizeFile_folderPath": Test_vtkNormalizeFile_folderPath
        Case "Test_vtkNormalizeFile_sameSizeBeforeThanAfter1": Test_vtkNormalizeFile_sameSizeBeforeThanAfter1
        Case "Test_vtkNormalizeFile_sameSizeBeforeThanAfter2": Test_vtkNormalizeFile_sameSizeBeforeThanAfter2
        Case "Test_vtkNormalizeFile_fileContent": Test_vtkNormalizeFile_fileContent
        Case "Test_vtkNormalizeFile_onlyOneFileInFolder": Test_vtkNormalizeFile_onlyOneFileInFolder
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

