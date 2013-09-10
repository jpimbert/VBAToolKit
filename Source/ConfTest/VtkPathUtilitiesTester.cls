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
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy.dummy"), "dummy", "dummy is the file name."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_fileNameNoExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy"), "dummy", "dummy is the file name."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_fileNameEmptyExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy."), "dummy", "dummy is the file name."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_pathExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy1\dummy.extension"), "dummy", "dummy is the file name."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_pathNoExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy1\dummy"), "dummy", "dummy is the file name."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_fileNameDoubleExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy.ext1.ext2"), "dummy.ext1", "dummy.ext1 is the file name."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_pathDoubleExtension()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy1\dummy.ext1.ext2"), "dummy.ext1", "dummy.ext1 is the file name."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_fileNameTripleExtensionEmptyThird()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy.ext1.ext2."), "dummy.ext1.ext2", "dummy.ext1 is the file name."
End Sub

Public Sub Test_vtkStripFilePathOrNameOfExtension_pathTripleExtensionEmptyThird()
    mAssert.Equals vtkStripFilePathOrNameOfExtension("dummy\dummy.ext1.ext2."), "dummy.ext1.ext2", "dummy.ext1 is the file name."
End Sub


Public Sub Test_vtkStripPathOrNameFromVtkExtension_fileNameNoVtkExtension()
    mAssert.Equals vtkStripPathOrNameFromVtkExtension("dummy.extension", "VTKEXT"), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameFromVtkExtension_pathNoVtkExtension()
    mAssert.Equals vtkStripPathOrNameFromVtkExtension("dummy1\dummy.extension", "VTKEXT"), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameFromVtkExtension_fileNameVtkExtension()
    mAssert.Equals vtkStripPathOrNameFromVtkExtension("dummy_VTKEXT.extension", "VTKEXT"), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameFromVtkExtension_pathVtkExtension()
    mAssert.Equals vtkStripPathOrNameFromVtkExtension("dummy1\dummy_VTKEXT.extension", "VTKEXT"), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameFromVtkExtension_nameVtkExtension()
    mAssert.Equals vtkStripPathOrNameFromVtkExtension("dummy_VTKEXT", "VTKEXT"), "dummy"
End Sub

Public Sub Test_vtkStripPathOrNameFromVtkExtension_fileNameEmptyVtkExtension()
    mAssert.Equals vtkStripPathOrNameFromVtkExtension("dummy_.extension", ""), "dummy_"
End Sub

Public Sub Test_vtkStripPathOrNameFromVtkExtension_pathEmptyVtkExtension()
    mAssert.Equals vtkStripPathOrNameFromVtkExtension("dummy1\dummy_.extension", ""), "dummy_"
End Sub

Public Sub Test_vtkStripPathOrNameFromVtkExtension_nameEmptyVtkExtension()
    mAssert.Equals vtkStripPathOrNameFromVtkExtension("dummy_", "VTKEXT"), "dummy_"
End Sub

Public Sub Test_vtkStripPathOrNameFromVtkExtension_fileNameVtkExtensionMistake()
    mAssert.Equals vtkStripPathOrNameFromVtkExtension("dummy_VTKEXT.extension", "VTKEXTMISTAKE"), "dummy_VTKEXT"
End Sub

Public Sub Test_vtkStripPathOrNameFromVtkExtension_pathVtkExtensionMistake()
    mAssert.Equals vtkStripPathOrNameFromVtkExtension("dummy1\dummy_VTKEXT.extension", "VTKEXTMISTAKE"), "dummy_VTKEXT"
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
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


