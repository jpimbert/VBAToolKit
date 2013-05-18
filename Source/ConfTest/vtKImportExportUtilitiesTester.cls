VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtKImportExportUtilitiesTester"
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
    Dim returnvalue As String
    
    Workbooks.Add.SaveAs (vtkTestPath & "\" & vtkTestProjectName & "_Test"), FileFormat:=(52) '52 is xlsm format
    returnvalue = vtkConfSheet()
End Sub

Private Sub ITestCase_TearDown()
Application.DisplayAlerts = False
    Workbooks(vtkTestProjectName & "_Test" & ".xlsm").Close
    'Delete Created WorkBook
    Kill vtkTestPath & "\" & vtkTestProjectName & "_Test" & ".xlsm"
Application.DisplayAlerts = True
End Sub
Public Sub TestConfSheetCorrectlyCreated()
'test that configuration sheet was created
Dim found As Boolean
Dim i As Integer
found = False
    For i = 1 To Sheets.Count
        If Sheets(i).name = "configurations" Then
            found = True
        End If
    Next
    
mAssert.Should found = True, "configurations sheet must be created"
End Sub
Public Sub TestofNotExistanceOfConfSheet()
'test that configuration sheet is correctly created
Dim found As Boolean
Dim i As Integer
    
    Sheets("configurations").Delete
    found = False
      For i = 1 To Sheets.Count
        If Sheets(i).name = "configurations" Then
            found = True
        End If
      Next
    
 mAssert.Should found = False, "configurations sheet must be deleted"
End Sub
Public Sub TestModuleNameRangeCorrectlyNammed()
    Dim returnvalue As String
    'return value = range letter
    returnvalue = vtkModuleNameRange()
mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(returnvalue & vtkFirstLine - 2), "Module Name", "range ModuleName must nammed correctly"
End Sub
Public Sub TestInformationRangeCorretlyNammed()
    Dim returnvalue As String
    'return value = range letter
    returnvalue = vtkInformationRange()
    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(returnvalue & vtkFirstLine - 3), "File Informations", "range File Informations must nammed correctly"
End Sub
Public Sub TestVbaUnitModuleNameCorrectlyExported()
  'test that cell how contain vbaunit module name before and after executing function "VtkInitializeExcelfileWithVbaUnitModuleName()"
    Dim returnvalue As Integer

    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine), "", "the cell must be empty , vbaunit module not yet exported"
    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + 17), "", "the range must be empty , vbaunit module not yet exported"

    returnvalue = VtkInitializeExcelfileWithVbaUnitModuleName()
     
    mAssert.Should ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine) <> "", "the cell must be empty , vbaunit module not yet exported"
    mAssert.Should ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + 17) <> "", "the range must be empty , vbaunit module not yet exported"

End Sub
Public Sub TestIsVbaUnitModule()
  'test function how return true if it's a vbaunit module and false if not
    Dim returnvalue As Boolean
    Dim returnvalue2 As Integer
 
    mAssert.Should vtkIsVbaUnit(ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine)) = False, "vbaunit list not yet created"
    mAssert.Should vtkIsVbaUnit(ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + 17)) = False, "vbaunit list not yet created"
 
    returnvalue2 = VtkInitializeExcelfileWithVbaUnitModuleName()
     
    mAssert.Should vtkIsVbaUnit(ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine)) = True, "vtkIsVbaUnit must return true the list is already created"
    mAssert.Should vtkIsVbaUnit(ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + 17)) = True, "vtkIsVbaUnit must return true the list is already created"
    mAssert.Should vtkIsVbaUnit(ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + 19)) = False, "this cell don't contain avalid vbaunit module"
 
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfSheetCorrectlyCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestofNotExistanceOfConfSheet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestModuleNameRangeCorrectlyNammed"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestInformationRangeCorretlyNammed"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestVbaUnitModuleNameCorrectlyExported"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestIsVbaUnitModule"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestConfSheetCorrectlyCreated": TestConfSheetCorrectlyCreated
        Case "TestofNotExistanceOfConfSheet": TestofNotExistanceOfConfSheet
        Case "TestModuleNameRangeCorrectlyNammed": TestModuleNameRangeCorrectlyNammed
        Case "TestInformationRangeCorretlyNammed": TestInformationRangeCorretlyNammed
        Case "TestVbaUnitModuleNameCorrectlyExported": TestVbaUnitModuleNameCorrectlyExported
        Case "TestIsVbaUnitModule": TestIsVbaUnitModule
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

