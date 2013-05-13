VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkVbaUnitTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
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
    'create xlsm file
      On Error Resume Next
      Dim Wb As Workbook
      Set Wb = Workbooks.Add
      'Save created project with xlsm extention
      Wb.SaveAs Filename:=(vtkTestPath & "\" & vtkTestProjectName & "_test"), FileFormat:=(52) '52 is xlsm format
      
      MkDir vtkTestPath & "\vbaunittest"
 
End Sub

Private Sub ITestCase_TearDown()
    On Error Resume Next
    Application.DisplayAlerts = False
    
    'Close Created WorkBook
    Workbooks(vtkTestProjectName & "_test" & ".xlsm").Close
    'Delete Created WorkBook
    Kill vtkTestPath & "\" & vtkTestProjectName & "_test" & ".xlsm"
    
    Kill vtkTestPath & "\vbaunittest\*"
    RmDir (vtkTestPath & "\vbaunittest")
    'delete all files in test folder
    
    Application.DisplayAlerts = True
End Sub

Public Sub TestVtkInitializeExcelfileWithVbaUnitModuleName()
'call function how intialize sheet with vbaunit module name
returnValue = VtkInitializeExcelfileWithVbaUnitModuleName(vtkTestProjectName & "_test")
For i = 1 To 19
'test that 19 first cell not empty
mAssert.Should Workbooks(vtkTestProjectName & "_test").Sheets(1).Range("A" & i) <> "", "cell must contain vbaunit module name"
Next
'test that 20 cell was empty
mAssert.Should Workbooks(vtkTestProjectName & "_test").Sheets(1).Range("A20") = "", "cell must be empty"

End Sub
Public Sub TestvtkExportVbaUnitModules()
Dim fso As New Scripting.FileSystemObject
'call function how intialize sheet with vbaunit module name
returnValue = vtkExportVbaUnitModules(vtkTestPath & "\vbaunitTest\", Left(ThisWorkbook.name, Len(ThisWorkbook.name) - 5))

mAssert.Should fso.GetFolder(vtkTestPath & "\vbaunitTest\").Files.Count = 18, "exported vbaunit file number must be equal to 18"

End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVtkInitializeExcelfileWithVbaUnitModuleName"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestvtkExportVbaUnitModules"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestVtkInitializeExcelfileWithVbaUnitModuleName": TestVtkInitializeExcelfileWithVbaUnitModuleName
        Case "TestvtkExportVbaUnitModules": TestvtkExportVbaUnitModules
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

