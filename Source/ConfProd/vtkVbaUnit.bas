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

Public Function vtkExportVbaUnitModules(DestinationPath As String, sourceworkbookname As String) As String
 Dim i As Integer
 Dim j As Integer
 Dim FullPAth As String
 Dim fso As New FileSystemObject
 Dim modulename As String
 
 For i = 1 To Workbooks(sourceworkbookname).VBProject.VBComponents.Count ' from 1 to number of modules,forms,... in the workbook
 For j = 2 To 19
 
 If Workbooks(sourceworkbookname & ".xlsm").VBProject.VBComponents.Item(i).name = Workbooks(sourceworkbookname & ".xlsm").Sheets(1).Range("A" & j) Then
   
   modulename = Workbooks(sourceworkbookname & ".xlsm").VBProject.VBComponents.Item(i).name
   
   Select Case Workbooks(sourceworkbookname).VBProject.VBComponents.Item(i).Type
     Case 1
        FullPAth = DestinationPath & modulename & ".bas" 'full path of file that will be created
     Case 2
        FullPAth = DestinationPath & modulename & ".cls" 'full path of file that will be created
   End Select
   
    If fso.FileExists(FullPAth) = False Then 'default function how verify if the file exist
       fso.CreateTextFile (FullPAth)        ' if the file don't exist we will create it
    End If
                   
    Workbooks(sourceworkbookname).VBProject.VBComponents.Item(i).Export (FullPAth) 'export module to the right folder
    Workbooks(sourceworkbookname & ".xlsm").Sheets(1).Range("b" & j) = "ok"
    Workbooks(sourceworkbookname & ".xlsm").Sheets(1).Range("c" & j) = FullPAth
 End If
 Next
 Next

End Function

