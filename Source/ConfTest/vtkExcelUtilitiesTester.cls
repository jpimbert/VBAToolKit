VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkExcelUtilitiesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    VBAToolKit.resetTestFolder
End Sub

Public Sub Test_vtkDefaultIsAddIn()
    mAssert.Should vtkDefaultIsAddIn("TestFile.xla"), "xla extension is Add-In"
    mAssert.Should vtkDefaultIsAddIn("TestFile.xlam"), "xlam extension is Add-In"
    mAssert.Should Not vtkDefaultIsAddIn("TestFile.xlsx"), "xlsx extension is not Add-In"
    mAssert.Should Not vtkDefaultIsAddIn("TestFile.xltx"), "xltx extension is not Add-In"
    mAssert.Should Not vtkDefaultIsAddIn("TestFile.xlsm"), "xlsm extension is not Add-In"
    mAssert.Should Not vtkDefaultIsAddIn("TestFile.xltm"), "xltm extension is not Add-In"
    mAssert.Should Not vtkDefaultIsAddIn("TestFile.xls"), "xls extension is not Add-In"
End Sub

Public Sub Test_vtkProtectProject()
    Dim Wb As Workbook, fullName As String, component As VBComponent
    ' Create a new workbook for test, and populate it with a code module
    Set Wb = vtkCreateExcelWorkbookForTestWithProjectName(projectName:="TestProject")
    Set component = Wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    component.CodeModule.AddFromString "Sub Test" & vbCrLf & "End Sub"
    ' Save, Close and Reopen the verify there is no protection
    fullName = Wb.fullName
    Wb.Close saveChanges:=True
    Set Wb = Workbooks.Open(fileName:=fullName)
    mAssert.Should Not Wb.VBProject.Protection, "Project must not be protected"
    ' Protect, then Save, close and reopen to verify there is a protection
    vtkProtectProject project:=Wb.VBProject, password:="aPassword"
    fullName = Wb.fullName
    Wb.Close saveChanges:=True
    Set Wb = Workbooks.Open(fileName:=fullName)
    mAssert.Should Wb.VBProject.Protection, "Project must be protected"
    ' Kill test workbook
    vtkCloseAndKillWorkbook Wb:=Wb
End Sub

Public Sub Test_vtkDefaultExcelExtension2003()
    vtkExcelVersionForTest = typical2003Version
    mAssert.Equals ".xls", vtkDefaultExcelExtension, "Default Excel 2003 extension"
End Sub

Public Sub Test_vtkDefaultExcelExtension2007()
    vtkExcelVersionForTest = typical2007Version
    mAssert.Equals ".xlsm", vtkDefaultExcelExtension, "Default Excel 2007 extension"
End Sub

Public Sub Test_vtkDefaultExcelAddInExtension2003()
    vtkExcelVersionForTest = typical2003Version
    mAssert.Equals ".xla", vtkDefaultExcelAddInExtension, "Default Excel 2003 Add-In extension"
End Sub

Public Sub Test_vtkDefaultExcelAddInExtension2007()
    vtkExcelVersionForTest = typical2007Version
    mAssert.Equals ".xlam", vtkDefaultExcelAddInExtension, "Default Excel 2007 Add-In extension"
End Sub

Public Sub Test_vtkExcelVersion_undefined()
    ' When the Excel version is not forced, The actal Excel version is returned
    mAssert.Equals vtkExcelVersion, Val(Application.Version), "Must be actual Excel version"
End Sub

Public Sub Test_vtkExcelVersion_defined()
    ' When the Excel version is forced, The forced Excel version is returned
    vtkExcelVersionForTest = 534
    mAssert.Equals vtkExcelVersion, 534, "Must be forced Excel version"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDefaultIsAddIn"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkProtectProject"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDefaultExcelExtension2003"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDefaultExcelExtension2007"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDefaultExcelAddInExtension2003"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkDefaultExcelAddInExtension2007"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExcelVersion_undefined"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Test_vtkExcelVersion_defined"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Test_vtkDefaultIsAddIn": Test_vtkDefaultIsAddIn
        Case "Test_vtkProtectProject": Test_vtkProtectProject
        Case "Test_vtkDefaultExcelExtension2003": Test_vtkDefaultExcelExtension2003
        Case "Test_vtkDefaultExcelExtension2007": Test_vtkDefaultExcelExtension2007
        Case "Test_vtkDefaultExcelAddInExtension2003": Test_vtkDefaultExcelAddInExtension2003
        Case "Test_vtkDefaultExcelAddInExtension2007": Test_vtkDefaultExcelAddInExtension2007
        Case "Test_vtkExcelVersion_undefined": Test_vtkExcelVersion_undefined
        Case "Test_vtkExcelVersion_defined": Test_vtkExcelVersion_defined
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

