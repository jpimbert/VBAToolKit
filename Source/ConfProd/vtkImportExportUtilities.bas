Attribute VB_Name = "vtkImportExportUtilities"
Option Explicit
'---------------------------------------------------------------------------------------
' Procedure : vtkConfSheet
' Author    : user
' Date      : 14/05/2013
' Purpose   : - create new sheet (if it not exist) that will contain table of parameters
'---------------------------------------------------------------------------------------
'
Public Function vtkConfSheet() As String
Dim sheetname
sheetname = "configurations"
On Error Resume Next
 Worksheets(sheetname).Select
 If Err <> 0 Then
 Worksheets.Add.name = sheetname
 End If
vtkConfSheet = sheetname
On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkModuleNameRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - return range name that contains list of modules
'             - write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleNameRange() As String
vtkModuleNameRange = "A"
ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine - 2) = "Module Name"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkModuleDevRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - return range name that contains list of path of developemnt configuration
'             - write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleDevRange() As String
vtkModuleDevRange = "B"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkModuleDeliveryRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - return range name that contains list of path of devivery configuration
'             - write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleDeliveryRange() As String
vtkModuleDeliveryRange = "C"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInformationRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - return range name that contains modules information
'             - write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkInformationRange() As String
vtkInformationRange = "D"
ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkInformationRange & vtkFirstLine - 3) = "File Informations"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkFirstLine
' Author    : user
' Date      : 13/05/2013
' Purpose   : - define the start line
'---------------------------------------------------------------------------------------
'
Public Function vtkFirstLine() As Integer
vtkFirstLine = 4
End Function
'---------------------------------------------------------------------------------------
' Procedure : VtkInitializeExcelfileWithVbaUnitModuleName
' Author    : user
' Date      : 09/05/2013
' Purpose   : - initialize ConfSheet with vbaunit module name
'             - Return the next first empty line number
'---------------------------------------------------------------------------------------
'
Public Function VtkInitializeExcelfileWithVbaUnitModuleName() As Integer

Dim tableofvbaunitname(17) As String
    
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
    tableofvbaunitname(0) = "VbaUnitMain"
Dim j As Integer
  For j = 0 To UBound(tableofvbaunitname) ' for j to table length
    ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & j + vtkFirstLine) = tableofvbaunitname(j)
    ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & j + vtkFirstLine).Interior.ColorIndex = 6
  Next

 VtkInitializeExcelfileWithVbaUnitModuleName = j + vtkFirstLine
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkIsVbaUnit
' Author    : user
' Date      : 17/05/2013
' Purpose   : - take name in parameter and verify if the module is a vbaunit module
'---------------------------------------------------------------------------------------
'
Public Function vtkIsVbaUnit(modulename As String) As Boolean
Dim i As Integer
Dim valinit As Integer
Dim valfin As Integer
    valinit = vtkFirstLine
    valfin = vtkFirstLine + 17
    vtkIsVbaUnit = False
 For i = vtkFirstLine To valfin
  If modulename = Range(vtkModuleNameRange & i) And modulename <> "" Then
     vtkIsVbaUnit = True
  Exit For
  End If
 Next
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkListAllModules
' Author    : user
' Date      : 17/05/2013
' Purpose   : - call VtkInitializeExcelfileWithVbaUnitModuleName and use his return value
'             - list all module of current project , verify that the module
'              is not a vbaunit and write his name in the range
'
'---------------------------------------------------------------------------------------
'
Public Function vtkListAllModules() As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim t As Integer

t = VtkInitializeExcelfileWithVbaUnitModuleName()
k = 0
  For i = 1 To ActiveWorkbook.VBProject.VBComponents.Count
    If vtkIsVbaUnit(ActiveWorkbook.VBProject.VBComponents.Item(i).name) = False Then
        ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & t + k) = ActiveWorkbook.VBProject.VBComponents.Item(i).name
        ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & t + k).Interior.ColorIndex = 8
        k = k + 1
    End If
  Next
vtkListAllModules = k
End Function
