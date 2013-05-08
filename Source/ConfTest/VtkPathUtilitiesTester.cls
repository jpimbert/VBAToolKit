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
    mAssert.Equals Mid(vtkPathOfCurrentProject, 2, 8), ":\Users\", "the install path must contain :\Users\ "
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathOfCurrentProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathToTestFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathToSourceFolder"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestPathfinishbyVBAToolKit"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestPathOfCurrentProject": TestPathOfCurrentProject
        Case "TestPathToTestFolder": TestPathToTestFolder
        Case "TestPathToSourceFolder": TestPathToSourceFolder
        Case "TestPathfinishbyVBAToolKit": TestPathfinishbyVBAToolKit
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


