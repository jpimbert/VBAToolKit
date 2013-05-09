Attribute VB_Name = "vtkVbaUnit"

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : VtkExistVbaUnit
' Author    : user
' Date      : 09/05/2013
' Purpose   : - initialize sheets with vbaunit module name
'---------------------------------------------------------------------------------------
'
Public Function VtkInitializeExcelfileWithVbaUnitModuleName(workbookname As String) As Boolean

Dim tableofvbaunitname(17) As String

tableofvbaunitname(0) = "VbaUnitMain"
tableofvbaunitname(1) = "IAssert"
tableofvbaunitname(2) = "IResultUser"
tableofvbaunitname(3) = "IRunManager"
tableofvbaunitname(4) = "ITest"
tableofvbaunitname(5) = "ITestCase"
tableofvbaunitname(6) = "ITestManager"
tableofvbaunitname(7) = "RunManager"
tableofvbaunitname(8) = "TestCaseManager"
tableofvbaunitname(9) = "TestClassLister"
tableofvbaunitname(10) = "TesterTemplate"
tableofvbaunitname(11) = "TestFailure"
tableofvbaunitname(12) = "TestResult"
tableofvbaunitname(13) = "TestRunner"
tableofvbaunitname(14) = "TestSuite"
tableofvbaunitname(15) = "TestSuiteManager"
tableofvbaunitname(16) = "AutoGen"
tableofvbaunitname(17) = "Assert"

Dim j As Integer
Workbooks(workbookname & ".xlsm").Sheets(1).Range("A1") = "VbaUnit Module Name"
Workbooks(workbookname & ".xlsm").Sheets(1).Range("A1").Interior.ColorIndex = 17
For j = 0 To UBound(tableofvbaunitname) ' for j to table length
Workbooks(workbookname & ".xlsm").Sheets(1).Range("A" & j + 2) = tableofvbaunitname(j)
Workbooks(workbookname & ".xlsm").Sheets(1).Range("A" & j + 2).Interior.ColorIndex = 15
Next
End Function
