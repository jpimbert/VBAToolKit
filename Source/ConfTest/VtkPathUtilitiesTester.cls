VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkPathUtilitiesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : VtkPathUtilitiesTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the vtkPathUtilities module functions
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
Private Const typical2003Version = 11
Private Const typical2007Version = 12

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
    vtkExcelVersionForTest = 0
End Sub

Public Sub TestPathToTestFolder()
    ' Verify the path of the current Test Folder
    mAssert.Equals vtkPathToTestFolder, vtkPathOfCurrentProject & "\Tests"
End Sub

Public Sub TestPathToSourceFolder()
    ' Verify the path of the current Source Folder
    mAssert.Equals vtkPathToSourceFolder, vtkPathOfCurrentProject & "\Source"
End Sub

Public Sub TestPathfinishbyVBAToolKit()
    ' Verify that the path finish by "\VBAToolKit"
    mAssert.Equals Right(vtkPathOfCurrentProject, 11), "\VBAToolKit", "the install path must finish by \VBAToolKit"
    ' Verify that the path contain ":\Users\"
    Dim bUser As Boolean, bDocument As Boolean
    bUser = (Mid(vtkPathOfCurrentProject, 2, 8) = ":\Users\")
    bDocument = (Mid(vtkPathOfCurrentProject, 2, 25) = ":\Documents and Settings\")
    mAssert.Should bUser Or bDocument, "the install path must contain :\Users\ or :\Documents\"
End Sub

Public Sub TestFileExtensionForStandardPath()
    mAssert.Equals vtkGetFileExtension("C:\Documents and Settings\User1\Documents\FileName.ext"), "ext", "The extension of this standard file is ext"
End Sub

Public Sub TestFileExtensionForFileName()
    mAssert.Equals vtkGetFileExtension("FileName.bas"), "bas", "The extension of this file name is bas"
End Sub

Public Sub TestFileExtensionWithTwoDots()
    mAssert.Equals vtkGetFileExtension("FileName.WithDot.xlsx"), "xlsx", "The extension of this file with a dot in name is xlsx"
End Sub

Public Sub TestFileExtensionWithEmptyExtension()
    mAssert.Equals vtkGetFileExtension("FileName."), "", "This file has a dot but no extension"
End Sub

Public Sub TestFileExtensionWithNoExtension()
    mAssert.Equals vtkGetFileExtension("FileName"), "", "This file has no extension"
End Sub

Public Sub TestFileExtensionWithJustAnExtension()
    mAssert.Equals vtkGetFileExtension(".FileName"), "FileName", "FileName is the extension"
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_fileNameExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy.dummy"), "dummy", "The extension must be stripped off."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_fileNameNoExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy"), "dummy", "No extension nor path : the file name must be returned as is."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_fileNameEmptyExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy."), "dummy", "The empty extension must be stripped off."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_pathExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummyFolder\dummy.extension"), "dummy", "The extension and the path must be stripped off"
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_pathNoExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummyFolder\dummy"), "dummy", "The path must be stripped off."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_fileNameDoubleExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy.ext1.ext2"), "dummy.ext1", "The last extension must be stripped off."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_pathDoubleExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummyFolder\dummy.ext1.ext2"), "dummy.ext1", "The last  extension and the path must be stripped off."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_fileNameTripleExtensionEmptyThird()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy.ext1.ext2."), "dummy.ext1.ext2", "The last extension (empty) must be stripped off."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_pathTripleExtensionEmptyThird()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy\dummy.ext1.ext2."), "dummy.ext1.ext2", "The last extension (empty) and the path must be stripped off."
End Sub


Public Sub Test_vtkStripPathOrNameOfVtkExtension_fileNameNoVtkExtension()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy.extension", "VTKEXT"), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_pathNoVtkExtension()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummyFolder\dummy.extension", "VTKEXT"), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_fileNameVtkExtension()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_VTKEXT.extension", "VTKEXT"), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_pathVtkExtension()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummyFolder\dummy_VTKEXT.extension", "VTKEXT"), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_nameVtkExtension()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_VTKEXT", "VTKEXT"), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension1()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_VTKEXT_OTHER.extension", "VTKEXT"), "dummy_VTKEXT_OTHER"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension2()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_OTHER_VTKEXT.extension", "VTKEXT"), "dummy_OTHER"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension3()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_VTKEXT_VTKEXT.extension", "VTKEXT"), "dummy_VTKEXT"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension1()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummyFolder\dummy_VTKEXT_OTHER.extension", "VTKEXT"), "dummy_VTKEXT_OTHER"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension2()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummyFolder\dummy_OTHER_VTKEXT.extension", "VTKEXT"), "dummy_OTHER"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension3()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummyFolder\dummy_VTKEXT_VTKEXT.extension", "VTKEXT"), "dummy_VTKEXT"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension1()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_VTKEXT_OTHER", "VTKEXT"), "dummy_VTKEXT_OTHER"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension2()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_OTHER_VTKEXT", "VTKEXT"), "dummy_OTHER"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension3()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_VTKEXT_VTKEXT", "VTKEXT"), "dummy_VTKEXT"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_fileNameEmptyVtkExtension()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_.extension", ""), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_pathEmptyVtkExtension()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummyFolder\dummy_.extension", ""), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_nameEmptyVtkExtension()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_", ""), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_fileNameVtkExtensionMistake()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummy_VTKEXT.extension", "VTKEXTMISTAKE"), "dummy_VTKEXT"
End Sub

Public Sub Test_vtkStripPathOrNameOfVtkExtension_pathVtkExtensionMistake()
    mAssert.Equals vtkStripPathOrNameOfVtkExtension("dummyFolder\dummy_VTKEXT.extension", "VTKEXTMISTAKE"), "dummy_VTKEXT"
End Sub

Public Sub Test_vtkIsPathAbsolute_ServerPath()
    Dim path As String
    path = "\\Server\dummyFolder\dummy_VTKEXT.extension"
    mAssert.Should vtkIsPathAbsolute(path), "Path " & path & " is absolute"
End Sub

Public Sub Test_vtkIsPathAbsolute_AbsDirPath()
    Dim path As String
    path = "\dummyFolder\dummy_VTKEXT.extension"
    mAssert.Should vtkIsPathAbsolute(path), "Path " & path & " is absolute"
End Sub

Public Sub Test_vtkIsPathAbsolute_UpcaseDrivePath()
    Dim path As String
    path = "C:\Server\dummyFolder\dummy_VTKEXT.extension"
    mAssert.Should vtkIsPathAbsolute(path), "Path " & path & " is absolute"
End Sub

Public Sub Test_vtkIsPathAbsolute_LowcaseDrivePath()
    Dim path As String
    path = "a:\Server\dummyFolder\dummy_VTKEXT.extension"
    mAssert.Should vtkIsPathAbsolute(path), "Path " & path & " is absolute"
End Sub

Public Sub Test_vtkIsPathAbsolute_RelativePath()
    Dim path As String
    path = "dummyFolder\dummy_VTKEXT.extension"
    mAssert.Should Not vtkIsPathAbsolute(path), "Path " & path & " is relative"
End Sub

Public Sub Test_vtkConvertGenericExcelPath_NoTemplate()
    ' Test the generic path conversion with no extension template
    Dim genericFilePath As String, expectedFilePath As String, actualFilePath As String
    
    genericFilePath = "dummyFolder\dummy_VTKEXT.extension"
    expectedFilePath = "dummyFolder\dummy_VTKEXT.extension"
    
    On Error Resume Next
    actualFilePath = vtkConvertGenericExcelPath(genericFilePath)
    mAssert.Equals Err.Number, 0, "Error returned by vtkConvertGenericExcelPath function"
    On Error GoTo 0
    
    mAssert.Equals actualFilePath, expectedFilePath, "Converted file path for generic path=" & genericFilePath
    
End Sub

Public Sub Test_vtkConvertGenericExcelPath_Irrelevant()
    ' Test the generic path conversion with irrelevant extension template
    Dim genericFilePath As String, expectedFilePath As String, actualFilePath As String
    
    genericFilePath = "dummyFolder\dummy_VTKEXT.<extension>"
    expectedFilePath = ""
    
    On Error Resume Next
    actualFilePath = vtkConvertGenericExcelPath(genericFilePath)
    mAssert.Equals Err.Number, VTK_WRONG_GENERIC_EXTENSION, "Error returned by vtkConvertGenericExcelPath function"
    On Error GoTo 0
    
    mAssert.Equals actualFilePath, expectedFilePath, "Converted file path for generic path=" & genericFilePath
    
End Sub

Public Sub Test_vtkConvertGenericExcelPath_XLA2003()
    ' Test the generic path conversion with <xla> extension template and Excel 2003
    Dim genericFilePath As String, expectedFilePath As String, actualFilePath As String
    
    genericFilePath = "dummyFolder\dummy_VTKEXT.<xla>"
    expectedFilePath = "dummyFolder\dummy_VTKEXT.xla"
    vtkExcelVersionForTest = typical2003Version
    
    On Error Resume Next
    actualFilePath = vtkConvertGenericExcelPath(genericFilePath)
    mAssert.Equals Err.Number, 0, "Error returned by vtkConvertGenericExcelPath function"
    On Error GoTo 0
    
    mAssert.Equals actualFilePath, expectedFilePath, "Converted file path for generic path=" & genericFilePath
    
End Sub

Public Sub Test_vtkConvertGenericExcelPath_XLA2007()
    ' Test the generic path conversion with <xla> extension template and Excel 2007
    Dim genericFilePath As String, expectedFilePath As String, actualFilePath As String
    
    genericFilePath = "dummyFolder\dummy_VTKEXT.<xla>"
    expectedFilePath = "dummyFolder\dummy_VTKEXT.xlam"
    vtkExcelVersionForTest = typical2007Version
    
    On Error Resume Next
    actualFilePath = vtkConvertGenericExcelPath(genericFilePath)
    mAssert.Equals Err.Number, 0, "Error returned by vtkConvertGenericExcelPath function"
    On Error GoTo 0
    
    mAssert.Equals actualFilePath, expectedFilePath, "Converted file path for generic path=" & genericFilePath
    
End Sub

Public Sub Test_vtkConvertGenericExcelPath_XLS2003()
    ' Test the generic path conversion with <xls> extension template and Excel 2003
    Dim genericFilePath As String, expectedFilePath As String, actualFilePath As String
    
    genericFilePath = "dummyFolder\dummy_VTKEXT.<xls>"
    expectedFilePath = "dummyFolder\dummy_VTKEXT.xls"
    vtkExcelVersionForTest = typical2003Version
    
    On Error Resume Next
    actualFilePath = vtkConvertGenericExcelPath(genericFilePath)
    mAssert.Equals Err.Number, 0, "Error returned by vtkConvertGenericExcelPath function"
    On Error GoTo 0
    
    mAssert.Equals actualFilePath, expectedFilePath, "Converted file path for generic path=" & genericFilePath
    
End Sub

Public Sub Test_vtkConvertGenericExcelPath_XLS2007()
    ' Test the generic path conversion with <xls> extension template and Excel 2007
    Dim genericFilePath As String, expectedFilePath As String, actualFilePath As String
    
    genericFilePath = "dummyFolder\dummy_VTKEXT.<xls>"
    expectedFilePath = "dummyFolder\dummy_VTKEXT.xlsm"
    vtkExcelVersionForTest = typical2007Version
    
    On Error Resume Next
    actualFilePath = vtkConvertGenericExcelPath(genericFilePath)
    mAssert.Equals Err.Number, 0, "Error returned by vtkConvertGenericExcelPath function"
    On Error GoTo 0
    
    mAssert.Equals actualFilePath, expectedFilePath, "Converted file path for generic path=" & genericFilePath
    
End Sub

Public Sub Test_vtkConvertGenericExcelPath_NoExtension()
    ' Test the generic path conversion without extension
    Dim genericFilePath As String, expectedFilePath As String, actualFilePath As String
    
    genericFilePath = "dummyFolder\dummy_VTKEXT"
    expectedFilePath = "dummyFolder\dummy_VTKEXT"
    
    On Error Resume Next
    actualFilePath = vtkConvertGenericExcelPath(genericFilePath)
    mAssert.Equals Err.Number, 0, "Error returned by vtkConvertGenericExcelPath function"
    On Error GoTo 0
    
    mAssert.Equals actualFilePath, expectedFilePath, "Converted file path for generic path=" & genericFilePath
    
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathToTestFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathToSourceFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathfinishbyVBAToolKit"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestFileExtensionForStandardPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestFileExtensionForFileName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestFileExtensionWithTwoDots"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestFileExtensionWithEmptyExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestFileExtensionWithNoExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestFileExtensionWithJustAnExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripFilePathOrNameOfExtension_fileNameExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripFilePathOrNameOfExtension_fileNameNoExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripFilePathOrNameOfExtension_fileNameEmptyExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripFilePathOrNameOfExtension_pathExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripFilePathOrNameOfExtension_pathNoExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripFilePathOrNameOfExtension_fileNameDoubleExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripFilePathOrNameOfExtension_pathDoubleExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripFilePathOrNameOfExtension_fileNameTripleExtensionEmptyThird"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripFilePathOrNameOfExtension_pathTripleExtensionEmptyThird"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_fileNameNoVtkExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_pathNoVtkExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_fileNameVtkExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_pathVtkExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_nameVtkExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension3"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension3"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension1"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension2"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension3"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_fileNameEmptyVtkExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_pathEmptyVtkExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_nameEmptyVtkExtension"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_fileNameVtkExtensionMistake"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkStripPathOrNameOfVtkExtension_pathVtkExtensionMistake"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkIsPathAbsolute_ServerPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkIsPathAbsolute_AbsDirPath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkIsPathAbsolute_UpcaseDrivePath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkIsPathAbsolute_LowcaseDrivePath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkIsPathAbsolute_RelativePath"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkConvertGenericExcelPath_NoTemplate"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkConvertGenericExcelPath_Irrelevant"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkConvertGenericExcelPath_XLA2003"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkConvertGenericExcelPath_XLA2007"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkConvertGenericExcelPath_XLS2003"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkConvertGenericExcelPath_XLS2007"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkConvertGenericExcelPath_NoExtension"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestPathToTestFolder": TestPathToTestFolder
        Case "TestPathToSourceFolder": TestPathToSourceFolder
        Case "TestPathfinishbyVBAToolKit": TestPathfinishbyVBAToolKit
        Case "TestFileExtensionForStandardPath": TestFileExtensionForStandardPath
        Case "TestFileExtensionForFileName": TestFileExtensionForFileName
        Case "TestFileExtensionWithTwoDots": TestFileExtensionWithTwoDots
        Case "TestFileExtensionWithEmptyExtension": TestFileExtensionWithEmptyExtension
        Case "TestFileExtensionWithNoExtension": TestFileExtensionWithNoExtension
        Case "TestFileExtensionWithJustAnExtension": TestFileExtensionWithJustAnExtension
        Case "Test_vtkStripFilePathOrNameOfExtension_fileNameExtension": Test_vtkStripFilePathOrNameOfExtension_fileNameExtension
        Case "Test_vtkStripFilePathOrNameOfExtension_fileNameNoExtension": Test_vtkStripFilePathOrNameOfExtension_fileNameNoExtension
        Case "Test_vtkStripFilePathOrNameOfExtension_fileNameEmptyExtension": Test_vtkStripFilePathOrNameOfExtension_fileNameEmptyExtension
        Case "Test_vtkStripFilePathOrNameOfExtension_pathExtension": Test_vtkStripFilePathOrNameOfExtension_pathExtension
        Case "Test_vtkStripFilePathOrNameOfExtension_pathNoExtension": Test_vtkStripFilePathOrNameOfExtension_pathNoExtension
        Case "Test_vtkStripFilePathOrNameOfExtension_fileNameDoubleExtension": Test_vtkStripFilePathOrNameOfExtension_fileNameDoubleExtension
        Case "Test_vtkStripFilePathOrNameOfExtension_pathDoubleExtension": Test_vtkStripFilePathOrNameOfExtension_pathDoubleExtension
        Case "Test_vtkStripFilePathOrNameOfExtension_fileNameTripleExtensionEmptyThird": Test_vtkStripFilePathOrNameOfExtension_fileNameTripleExtensionEmptyThird
        Case "Test_vtkStripFilePathOrNameOfExtension_pathTripleExtensionEmptyThird": Test_vtkStripFilePathOrNameOfExtension_pathTripleExtensionEmptyThird
        Case "Test_vtkStripPathOrNameOfVtkExtension_fileNameNoVtkExtension": Test_vtkStripPathOrNameOfVtkExtension_fileNameNoVtkExtension
        Case "Test_vtkStripPathOrNameOfVtkExtension_pathNoVtkExtension": Test_vtkStripPathOrNameOfVtkExtension_pathNoVtkExtension
        Case "Test_vtkStripPathOrNameOfVtkExtension_fileNameVtkExtension": Test_vtkStripPathOrNameOfVtkExtension_fileNameVtkExtension
        Case "Test_vtkStripPathOrNameOfVtkExtension_pathVtkExtension": Test_vtkStripPathOrNameOfVtkExtension_pathVtkExtension
        Case "Test_vtkStripPathOrNameOfVtkExtension_nameVtkExtension": Test_vtkStripPathOrNameOfVtkExtension_nameVtkExtension
        Case "Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension1": Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension1
        Case "Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension2": Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension2
        Case "Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension3": Test_vtkStripPathOrNameOfVtkExtension_fileNameDoubleVtkExtension3
        Case "Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension1": Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension1
        Case "Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension2": Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension2
        Case "Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension3": Test_vtkStripPathOrNameOfVtkExtension_pathDoubleVtkExtension3
        Case "Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension1": Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension1
        Case "Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension2": Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension2
        Case "Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension3": Test_vtkStripPathOrNameOfVtkExtension_nameDoubleVtkExtension3
        Case "Test_vtkStripPathOrNameOfVtkExtension_fileNameEmptyVtkExtension": Test_vtkStripPathOrNameOfVtkExtension_fileNameEmptyVtkExtension
        Case "Test_vtkStripPathOrNameOfVtkExtension_pathEmptyVtkExtension": Test_vtkStripPathOrNameOfVtkExtension_pathEmptyVtkExtension
        Case "Test_vtkStripPathOrNameOfVtkExtension_nameEmptyVtkExtension": Test_vtkStripPathOrNameOfVtkExtension_nameEmptyVtkExtension
        Case "Test_vtkStripPathOrNameOfVtkExtension_fileNameVtkExtensionMistake": Test_vtkStripPathOrNameOfVtkExtension_fileNameVtkExtensionMistake
        Case "Test_vtkStripPathOrNameOfVtkExtension_pathVtkExtensionMistake": Test_vtkStripPathOrNameOfVtkExtension_pathVtkExtensionMistake
        Case "Test_vtkIsPathAbsolute_ServerPath": Test_vtkIsPathAbsolute_ServerPath
        Case "Test_vtkIsPathAbsolute_AbsDirPath": Test_vtkIsPathAbsolute_AbsDirPath
        Case "Test_vtkIsPathAbsolute_UpcaseDrivePath": Test_vtkIsPathAbsolute_UpcaseDrivePath
        Case "Test_vtkIsPathAbsolute_LowcaseDrivePath": Test_vtkIsPathAbsolute_LowcaseDrivePath
        Case "Test_vtkIsPathAbsolute_RelativePath": Test_vtkIsPathAbsolute_RelativePath
        Case "Test_vtkConvertGenericExcelPath_NoTemplate": Test_vtkConvertGenericExcelPath_NoTemplate
        Case "Test_vtkConvertGenericExcelPath_Irrelevant": Test_vtkConvertGenericExcelPath_Irrelevant
        Case "Test_vtkConvertGenericExcelPath_XLA2003": Test_vtkConvertGenericExcelPath_XLA2003
        Case "Test_vtkConvertGenericExcelPath_XLA2007": Test_vtkConvertGenericExcelPath_XLA2007
        Case "Test_vtkConvertGenericExcelPath_XLS2003": Test_vtkConvertGenericExcelPath_XLS2003
        Case "Test_vtkConvertGenericExcelPath_XLS2007": Test_vtkConvertGenericExcelPath_XLS2007
        Case "Test_vtkConvertGenericExcelPath_NoExtension": Test_vtkConvertGenericExcelPath_NoExtension
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


