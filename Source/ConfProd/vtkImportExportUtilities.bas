Attribute VB_Name = "vtkImportExportUtilities"
Option Explicit



Private newWorkbook As Workbook         ' New Workbook created for each test
Private newWorkbookName As String
Private ConfManager As vtkConfigurationManager   ' Configuration Manager for the new workbook
Public DevWbFullName As String
Public DelivwbFullPAth As String
Public DelivwbName As String
Public DevwbFullPAth As String
Public DevwbName As String


'---------------------------------------------------------------------------------------
' Procedure : VtkInitilizeSheet
' Author    : user
' Date      : 03/06/2013
' Purpose   : - function how call function to initilize worksheet
'---------------------------------------------------------------------------------------
'
Public Function VtkInitilizeSheet()
    Dim i As Integer
    Dim j As Integer
    Dim a

    Set ConfManager = vtkConfigurationManagerForWorkbook(DevWbFullName)

    ConfManager.setConfigurationPathWithNumber n:=1, Path:=DelivwbFullPAth
    ConfManager.setConfigurationPathWithNumber n:=2, Path:=DevwbFullPAth
    i = VtkInitializeExcelfileWithVbaUnitModule()
    j = vtkListAllModules()

End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkFirstLine
' Author    : user
' Date      : 13/05/2013
' Purpose   : - define the start line
'---------------------------------------------------------------------------------------
'
Public Function vtkFirstLine() As Integer
vtkFirstLine = 3
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkModuleNameRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - define module column name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleNameRange() As String
vtkModuleNameRange = "A"
End Function

'---------------------------------------------------------------------------------------
' Procedure : VtkInitializeExcelfileWithVbaUnitModuleName
' Author    : user
' Date      : 09/05/2013
' Purpose   : - initialize ConfSheet with vbaunit module name
'             - Return the next first empty line number
'             - export vbaunit module from vbatoolkit workbook to the right path of the created project
'             - import vbaunitmodule from the created project to the new workbook
'---------------------------------------------------------------------------------------
'
Public Function VtkInitializeExcelfileWithVbaUnitModule() As Integer


    Dim retval
    Dim terval2
    Dim Path As String
    Dim fullpath As String
    Dim j As Integer
    Dim fso As New FileSystemObject

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

 Path = fso.GetParentFolderName(ActiveWorkbook.Path)
  
  For j = 0 To UBound(tableofvbaunitname) ' for j to table length
      
    'add module to the confsheet
    retval = ConfManager.addModule(module:=tableofvbaunitname(j))
     
     'set the rightpath for the module
       If tableofvbaunitname(j) = "VbaUnitMain" Then
             fullpath = Path & "\Source\VbaUnit\" & tableofvbaunitname(j) & ".bas"  'full path of file that will be created
             ConfManager.setModulePathWithNumber Path:=fullpath, numModule:=j, numConfiguration:=2
          Else
             fullpath = Path & "\Source\VbaUnit\" & tableofvbaunitname(j) & ".cls"  'full path of file that will be created
             ConfManager.setModulePathWithNumber Path:=fullpath, numModule:=j, numConfiguration:=2
       End If
   'export module from source workbook to the created project folder
   Workbooks(ThisWorkbook.name).VBProject.VBComponents(tableofvbaunitname(j)).Export (fullpath)
   'import module from the new project folder to the new workbook
   Workbooks(Dir(ConfManager.getConfigurationPathWithNumber(2))).VBProject.VBComponents.Import fullpath
 
 Next
 
  VtkInitializeExcelfileWithVbaUnitModule = vtkFirstLine + j
End Function

'---------------------------------------------------------------------------------------
' Procedure : ModuleNameAlreadyExistInSheet
' Author    : user
' Date      : 17/05/2013
' Purpose   : - take name in parameter and verify if the module is a vbaunit module
'---------------------------------------------------------------------------------------
'
Public Function ModuleNameAlreadyExistInSheet(modulename As String) As Boolean
Dim i As Integer

    ModuleNameAlreadyExistInSheet = False
 For i = vtkFirstLine To ConfManager.moduleCount
  If modulename = Range(vtkModuleNameRange & i) Then
     ModuleNameAlreadyExistInSheet = True
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
'             -!! must verify deleted method from workbook
'---------------------------------------------------------------------------------------
'
Public Function vtkListAllModules() As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim retval
k = 0
  For i = 1 To ActiveWorkbook.VBProject.VBComponents.Count
    If ModuleNameAlreadyExistInSheet(ActiveWorkbook.VBProject.VBComponents.Item(i).name) = False Then
      retval = ConfManager.addModule(ActiveWorkbook.VBProject.VBComponents.Item(i).name)
        k = k + 1
    End If
  Next
vtkListAllModules = k
End Function
