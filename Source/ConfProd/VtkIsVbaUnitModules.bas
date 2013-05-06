Attribute VB_Name = "VtkIsVbaUnitModules"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : existvbaunit
' Author    : user
' Date      : 22/04/2013
' Purpose   : - initialize table with vbaunit modules name
'             - return true if var exist on table
'---------------------------------------------------------------------------------------
'
Public Function VtkExistVbaUnit(nameofmodule As String) As Boolean

Dim tableofvbaunitname(17) As String
tableofvbaunitname(0) = "AutoGen"
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
tableofvbaunitname(16) = "VbaUnitMain"
tableofvbaunitname(17) = "Assert"

VtkExistVbaUnit = False ' initialize function by false
Dim j As Integer
For j = 0 To UBound(tableofvbaunitname) ' for j to table length

    If UCase(nameofmodule) = UCase(tableofvbaunitname(j)) Then 'compare majuscule of tablecontent with majuscule of parameters
        VtkExistVbaUnit = True      ' return true if var exist
    End If
Next
End Function
