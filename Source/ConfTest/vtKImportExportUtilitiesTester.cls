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


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestConfSheetCorrectlyCreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestofNotExistanceOfConfSheet"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestModuleNameRangeCorrectlyNammed"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestInformationRangeCorretlyNammed"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestConfSheetCorrectlyCreated": TestConfSheetCorrectlyCreated
        Case "TestofNotExistanceOfConfSheet": TestofNotExistanceOfConfSheet
        Case "TestModuleNameRangeCorrectlyNammed": TestModuleNameRangeCorrectlyNammed
        Case "TestInformationRangeCorretlyNammed": TestInformationRangeCorretlyNammed
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

