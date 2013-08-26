VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkCreateTreeFolderTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkCreateTreeFolderTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the vtkCeateTreeFolder function
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
'   Delete Folder Tree
    vtkCleanFolder vtkTestPath
End Sub

Public Sub TestCreateTreeFolderDoesNothingWhenPathDoesntExist()
    ' The Create Tree Folder function must does nothing and return error when path doesn't exist
    Dim InexistentPath As String
    Dim returnvalue As Long
    InexistentPath = vtkTestPath & "\InexistentFolder"
    returnvalue = vtkCreateTreeFolder(rootPath:=InexistentPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(InexistentPath & "\" & vtkTestProjectName, vbDirectory) = "", "The project main folder must not exist"
    mAssert.Equals returnvalue, 76, "The project creation must fail"
End Sub

Public Sub TestCreateTreeFolderAlreadyExisting()
    ' The Create Tree Folder function must does nothing when project already exists
    Dim returnvalue As Long
    MkDir vtkTestPath & "\" & vtkTestProjectName 'must respect order mkdir before call vtkcreateproject
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName, vbDirectory) <> "", "The existing project main folder must exist"
    mAssert.Equals returnvalue, 75, "The project creation must fail"
End Sub

Public Sub TestMainFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named like the project in the path
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName, vbDirectory) <> "", "The project main folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestProjectFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named Project
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Project", vbDirectory) <> "", "The Project folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestSourceFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named Source
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Source", vbDirectory) <> "", "The Source folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestTestsFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named Tests
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Tests", vbDirectory) <> "", "The Tests folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestConfProdFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named ConfProd
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\ConfProd", vbDirectory) <> "", "source\ConfProd folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestConfTestFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named ConfProd
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\ConfTest", vbDirectory) <> "", "The source\ConfTest folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestVbaUnitFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named ConfProd
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\VbaUnit", vbDirectory) <> "", "The source\VbaUnit folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestGitLogIsCreated()
    ' The Create Tree Folder function must create a new folder named GitLog
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\GitLog", vbDirectory) <> "", "The ..\GitLog folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Public Sub TestDeliveryFolderIsCreated()
    ' The Create Tree Folder function must create a new folder named Delivery
    Dim returnvalue As Long
    returnvalue = vtkCreateTreeFolder(rootPath:=vtkTestPath & "\" & vtkTestProjectName)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Delivery", vbDirectory) <> "", "The \Delivery folder must exist"
    mAssert.Equals returnvalue, 0, "The project creation must succeed"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateTreeFolderDoesNothingWhenPathDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateTreeFolderAlreadyExisting"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestMainFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestProjectFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestSourceFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestTestsFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfProdFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfTestFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVbaUnitFolderIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestGitLogIsCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestDeliveryFolderIsCreated"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestCreateTreeFolderDoesNothingWhenPathDoesntExist": TestCreateTreeFolderDoesNothingWhenPathDoesntExist
        Case "TestCreateTreeFolderAlreadyExisting": TestCreateTreeFolderAlreadyExisting
        Case "TestMainFolderIsCreated": TestMainFolderIsCreated
        Case "TestProjectFolderIsCreated": TestProjectFolderIsCreated
        Case "TestSourceFolderIsCreated": TestSourceFolderIsCreated
        Case "TestTestsFolderIsCreated": TestTestsFolderIsCreated
        Case "TestConfProdFolderIsCreated": TestConfProdFolderIsCreated
        Case "TestConfTestFolderIsCreated": TestConfTestFolderIsCreated
        Case "TestVbaUnitFolderIsCreated": TestVbaUnitFolderIsCreated
        Case "TestGitLogIsCreated": TestGitLogIsCreated
        Case "TestDeliveryFolderIsCreated": TestDeliveryFolderIsCreated
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


