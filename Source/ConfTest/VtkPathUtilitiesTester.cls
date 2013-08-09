VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VtkPathUtilitiesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Sub TestPathOfCurrentProject()
    ' Verify the path of the current project is where VBATollKit is installed
    mAssert.Equals vtkPathOfCurrentProject, vtkInstallPath
End Sub

Public Sub TestPathToTestFolder()
    ' Verify the path of the current Test Folder
    mAssert.Equals vtkPathToTestFolder, vtkInstallPath & "\Tests"
End Sub

Public Sub TestPathToSourceFolder()
    ' Verify the path of the current Source Folder
    mAssert.Equals vtkPathToSourceFolder, vtkInstallPath & "\Source"
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

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathOfCurrentProject"
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
        Case "TestPathOfCurrentProject": TestPathOfCurrentProject
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


