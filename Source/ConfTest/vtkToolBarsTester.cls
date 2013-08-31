VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkToolBarsTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkToolBarsTester
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Test the vtkToolBars module
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

Private Const toolBarName = "VBAToolKit_DEV"

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
    vtkTestCommandBarButtonClickedReset
End Sub

Private Sub ITestCase_TearDown()
    vtkDeleteToolbars   ' The test toolbars have to be deleted after Unit Test ran
End Sub

'
'---------------------------------------------------------------------------------------
' Toolbar tests
'---------------------------------------------------------------------------------------
'
Public Sub TestExcelToolBarCreation()
    '   Verify that the Excel toolbar exists after creation
    Dim barE As CommandBar, barV As CommandBar
    
    vtkCreateToolbars vbeToolbar:=False
    
    On Error Resume Next
    Set barE = Application.CommandBars(toolBarName)
    Set barV = Application.VBE.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should Not barE Is Nothing, "The Excel Toolbar must exist"
    mAssert.Should barV Is Nothing, "The VBE Toolbar must not exist"
End Sub

Public Sub TestVBEToolBarCreation()
    '   Verify that the VBE toolbar exists after creation
    Dim barE As CommandBar, barV As CommandBar
    
    vtkCreateToolbars excToolbar:=False
    
    On Error Resume Next
    Set barE = Application.CommandBars(toolBarName)
    Set barV = Application.VBE.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should barE Is Nothing, "The Excel Toolbar must not exist"
    mAssert.Should Not barV Is Nothing, "The VBE Toolbar must exist"

End Sub

Public Sub TestBothToolBarCreation()
    '   Verify that the VBE and Excel toolbars exist after creation
    Dim barE As CommandBar, barV As CommandBar
    
    vtkCreateToolbars
    
    On Error Resume Next
    Set barE = Application.CommandBars(toolBarName)
    Set barV = Application.VBE.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should Not barE Is Nothing, "The Excel Toolbar must exist"
    mAssert.Should Not barV Is Nothing, "The VBE Toolbar must exist"

End Sub

Public Sub TestExcelToolBarDeletion()
    '   Verify that the Excel toolbar doesn't exist after deletion
    Dim bar As CommandBar
    
    vtkCreateToolbars
    vtkDeleteToolbars
    
    On Error Resume Next
    Set bar = Application.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should bar Is Nothing, "The Excel Toolbar must not exist"
End Sub

Public Sub TestVBEToolBarDeletion()
    '   Verify that the VBE toolbar doesn't exist after deletion
    Dim bar As CommandBar
    
    vtkCreateToolbars
    vtkDeleteToolbars
    
    On Error Resume Next
    Set bar = Application.VBE.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should bar Is Nothing, "The VBE Toolbar must not exist"
End Sub

'
' Tests pour l'ajout de boutons
'   - vérifier que le bouton est ajouté (Excel, VBE)- il faudrait un ID
'   - vérifier que les boutons sont détruits avec la barre d'outils (Excel, VBE)
'   - vérifier que le bouton est actif (Excel, VBE)
'
'---------------------------------------------------------------------------------------
' Button tests
'---------------------------------------------------------------------------------------
'
Public Sub Test_ExcelToolBarButtonAdd()
    '   Verify that a ToolBarButton can be created in the Excel Toolbar
    Dim bar As CommandBar, button As CommandBarButton, realCaption As String

    vtkCreateEmptyToolbars   ' create empty toolbar
    vtkCreateToolbarButton caption:="Test button", helpText:="", faceId:=1, action:="vtkTestCommandBarButtonClicked", vbeToolbar:=False
    
   On Error Resume Next
    Set bar = Application.CommandBars(toolBarName)
    Set button = bar.Controls(1)
    realCaption = button.caption
   On Error GoTo 0
    
    mAssert.Equals realCaption, "Test button", "A Button must be created in the Excel ToolBar"
End Sub

Public Sub Test_VBEToolBarButtonAdd()
    '   Verify that a ToolBarButton can be created in the VBE Toolbar
    Dim bar As CommandBar, button As CommandBarButton, realCaption As String

    vtkCreateEmptyToolbars   ' create empty toolbar
    vtkCreateToolbarButton caption:="Test button", helpText:="", faceId:=1, action:="vtkTestCommandBarButtonClicked", excToolbar:=False
    
   On Error Resume Next
    Set bar = Application.VBE.CommandBars(toolBarName)
    Set button = bar.Controls(1)
    realCaption = button.caption
   On Error GoTo 0
    
    mAssert.Equals realCaption, "Test button", "A Button must be created in the Excel ToolBar"
End Sub

Public Sub Test_ExcelToolBarButtonDelete()
    '   Verify that a ToolBarButton is deleted when the Excel Toolbar is deleted
    Dim bar As CommandBar, button As CommandBarButton, realCaption As String, errNum As Long

    vtkCreateEmptyToolbars   ' create empty toolbar
    vtkCreateToolbarButton caption:="Test button", helpText:="", faceId:=1, action:="vtkTestCommandBarButtonClicked"
    Set bar = Application.CommandBars(toolBarName)
    Set button = bar.Controls(1)
    
   On Error Resume Next
    vtkDeleteToolbars
    realCaption = button.caption
    errNum = err.Number
   On Error GoTo 0
    
    mAssert.Equals errNum, 424, "No Button must be present if the Excel ToolBar is deleted"
End Sub

Public Sub Test_VBEToolBarButtonDelete()
    '   Verify that a ToolBarButton is deleted then the VBE Toolbar is deleted
    Dim bar As CommandBar, button As CommandBarButton, realCaption As String, errNum As Long

    vtkCreateEmptyToolbars   ' create empty toolbar
    vtkCreateToolbarButton caption:="Test button", helpText:="", faceId:=1, action:="vtkTestCommandBarButtonClicked"
    Set bar = Application.VBE.CommandBars(toolBarName)
    Set button = bar.Controls(1)
    
   On Error Resume Next
    vtkDeleteToolbars
    realCaption = button.caption
    errNum = err.Number
   On Error GoTo 0
    
    mAssert.Equals errNum, 424, "No Button must be present if the VBE ToolBar is deleted"
End Sub

Public Sub Test_ExcelToolBarButtonClick()
    '   Verify that a ToolBarButton activation call the dedicated Sub
    Dim button As CommandBarButton

    vtkCreateEmptyToolbars   ' create empty toolbar
    vtkCreateToolbarButton caption:="Test button", helpText:="", faceId:=1, action:="vtkTestCommandBarButtonClicked"
    Set button = Application.CommandBars(toolBarName).Controls(1)
        
    button.Execute
    
    mAssert.Should vtkIsTestCommandBarButtonClicked, "The Excel Toolbar button click must call the dedicated Sub"
End Sub

Public Sub Test_VBEToolBarButtonClick()
    '   Verify that a ToolBarButton activation call the dedicated Sub
    Dim button As CommandBarButton

    vtkCreateEmptyToolbars   ' create empty toolbar
    vtkCreateToolbarButton caption:="Test button", helpText:="", faceId:=1, action:="vtkTestCommandBarButtonClicked"
    Set button = Application.VBE.CommandBars(toolBarName).Controls(1)
        
    button.Execute
    
    mAssert.Should vtkIsTestCommandBarButtonClicked, "The VBE Toolbar button click must call the dedicated Sub"
End Sub

'
'---------------------------------------------------------------------------------------
' AddIn activation tests
'---------------------------------------------------------------------------------------
'
Public Sub TestAddInInstall()
    '   Verify that the Add-In installation creates the ToolBars
    Dim barE As CommandBar, barV As CommandBar
    
    Application.Run Macro:="VBAToolKit_DEV.ThisWorkBook.Workbook_AddinInstall"
    
    On Error Resume Next
    Set barE = Application.CommandBars(toolBarName)
    Set barV = Application.VBE.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should Not barE Is Nothing, "The Excel Toolbar must exist"
    mAssert.Should Not barV Is Nothing, "The VBE Toolbar must exist"
End Sub

Public Sub TestAddInUnInstall()
    '   Verify that the Add-In Uninstallation deletes the ToolBars
    Dim barE As CommandBar, barV As CommandBar
    
    vtkCreateToolbars
    Application.Run Macro:="VBAToolKit_DEV.ThisWorkBook.Workbook_AddinUninstall"
    
    On Error Resume Next
    Set barE = Application.CommandBars(toolBarName)
    Set barV = Application.VBE.CommandBars(toolBarName)
    On Error GoTo 0
    
    mAssert.Should barE Is Nothing, "The Excel Toolbar must not exist"
    mAssert.Should barV Is Nothing, "The VBE Toolbar must not exist"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestExcelToolBarCreation"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVBEToolBarCreation"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestBothToolBarCreation"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestExcelToolBarDeletion"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVBEToolBarDeletion"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ExcelToolBarButtonAdd"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_VBEToolBarButtonAdd"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ExcelToolBarButtonDelete"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_VBEToolBarButtonDelete"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_ExcelToolBarButtonClick"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_VBEToolBarButtonClick"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddInInstall"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestAddInUnInstall"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestExcelToolBarCreation": TestExcelToolBarCreation
        Case "TestVBEToolBarCreation": TestVBEToolBarCreation
        Case "TestBothToolBarCreation": TestBothToolBarCreation
        Case "TestExcelToolBarDeletion": TestExcelToolBarDeletion
        Case "TestVBEToolBarDeletion": TestVBEToolBarDeletion
        Case "Test_ExcelToolBarButtonAdd": Test_ExcelToolBarButtonAdd
        Case "Test_VBEToolBarButtonAdd": Test_VBEToolBarButtonAdd
        Case "Test_ExcelToolBarButtonDelete": Test_ExcelToolBarButtonDelete
        Case "Test_VBEToolBarButtonDelete": Test_VBEToolBarButtonDelete
        Case "Test_ExcelToolBarButtonClick": Test_ExcelToolBarButtonClick
        Case "Test_VBEToolBarButtonClick": Test_VBEToolBarButtonClick
        Case "TestAddInInstall": TestAddInInstall
        Case "TestAddInUnInstall": TestAddInUnInstall
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

