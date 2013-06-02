Attribute VB_Name = "vtkImportExportUtilities"
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : vtkModuleNameRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - return range name that contains list of modules
'
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleNameRange() As String
vtkModuleNameRange = "A"
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
End Function
'---------------------------------------------------------------------------------------
' Procedure : vtkModuleInformationsRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - return range name that contains modules information
'
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleInformationsRange() As String
vtkModuleInformationsRange = "E"
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
' Procedure : VtkInitializeExcelfileWithVbaUnitModuleName
' Author    : user
' Date      : 09/05/2013
' Purpose   : - initialize ConfSheet with vbaunit module name
'             - Return the next first empty line number
'             - vbaunit path line will be colored differently , to export it only one time when we initilize worksheet
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
    ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & j + vtkFirstLine).Interior.ColorIndex = 3
  Next
  VtkInitializeExcelfileWithVbaUnitModuleName = vtkFirstLine + j
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

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateModuleFile
' Author    : user
' Date      : 17/05/2013
' Purpose   : - this function allow to create a file
'             - return message contain informations: time , file created or replaced
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateModuleFile(fullpath As String) As String

Dim fso As New FileSystemObject

If fso.FileExists(fullpath) = False Then
    fso.CreateTextFile (fullpath)
vtkCreateModuleFile = "File created successfully at" & Now
Else
vtkCreateModuleFile = "File last update at" & Now
End If
End Function



'---------------------------------------------------------------------------------------
' Procedure : vtkExportModule
' Author    : user
' Date      : 14/05/2013
' Purpose   : - function take modulename , and line number , and workbookSource Name
'             - create files of modules if they don't exist ,or update it
'             - export module to the right folders  (documents , worksheets)
'             - write creation file informations
'             - write exported file location
'
'  if "vbaunitclass" then
'       if vbaUnitMain then ===================>path= vbaunit ".bas"
'       else                ===================>path= vbaunit ".cls"
'       endif
'  else
'     case module.type
'
'       1.module ,to ===========================>path= confprod ".BAS"
'       2.classmodule, if---nameTester to ======>path= ConfTest ".CLS"
'                      else ====================>path= ConfProd ".CLS"
'       3.Form   ,to ===========================>path= confprod ".FRM"
'     sheet ,worksheet, workbook ===============> do nothing
'  endif
'  vtkCreateModuleFile(path)
'  sheet.range = path
'---------------------------------------------------------------------------------------
'
Public Function vtkExportModule(modulename As String, lineNumber As Integer, sourceworkbook As String) As String

 Dim fullpath As String
 Dim path As String
 Dim MsgCreationFile As String
 Dim Test As String
 Dim DevPath As String
 Dim DelivPath As String
 Dim color As Integer
 color = 2
 Dim fso As New FileSystemObject
 path = fso.GetParentFolderName(ActiveWorkbook.path)
 
 
    If vtkIsVbaUnit(modulename) = True Then
          If modulename = "VbaUnitMain" Then
                fullpath = path & "\Source\VbaUnit\" & modulename & ".bas"  'full path of file that will be created
                DevPath = fullpath
                DelivPath = ""
                color = 3
          Else
                fullpath = path & "\Source\VbaUnit\" & modulename & ".cls"  'full path of file that will be created
                DevPath = fullpath
                DelivPath = ""
                color = 3
          End If
    Else
        
        On Error Resume Next
    
    
    Select Case Workbooks(sourceworkbook).VBProject.VBComponents(modulename).Type
          
        Case 1 '1module : export to confprod
         
           fullpath = path & "\Source\ConfProd\" & modulename & ".bas"  'full path of file that will be created
            DevPath = fullpath
            DelivPath = fullpath
          
         
        Case 2 '2 class module : export to ConfTest or ConfProd
         
            If Right(modulename, 6) Like "Tester" Then ' verify if modulename end is like Tester
                    
                ' This Document is a test module export to confTest
                fullpath = path & "\Source\ConfTest\" & modulename & ".CLS"
                DevPath = fullpath
                DelivPath = ""
                color = 3
            Else
    
                'the document is a classmodule export to confprod
                fullpath = path & "\Source\ConfProd\" & modulename & ".CLS"
                DevPath = fullpath
                DelivPath = fullpath
            End If
        Case 3 '3 forms
        
                'the document is a classmodule export to confprod
                fullpath = path & "\Source\ConfProd\" & modulename & ".FRM"
                DevPath = fullpath
                DelivPath = fullpath
                
        Case 100 'excel sheets , we will not export them for the moment
                DevPath = ""
                DelivPath = ""
                color = 3
                ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & lineNumber).Interior.ColorIndex = color
                
        Case Else 'normally we haven't other type but if we find another type we will export it to main project folder
                DevPath = ""
                DelivPath = ""
                color = 3
                ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & lineNumber).Interior.ColorIndex = color
          Exit Function
          
      End Select
    End If

   MsgCreationFile = vtkCreateModuleFile(fullpath)
   Workbooks(sourceworkbook).VBProject.VBComponents(modulename).Export (fullpath) 'export module to the right folder

   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & lineNumber) = DevPath
   
   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & lineNumber) = DelivPath
   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & lineNumber).Interior.ColorIndex = color
   
   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkInformationRange & lineNumber) = MsgCreationFile
 
   On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkExportAll
' Author    : user
' Date      : 16/05/2013
' Purpose   : - call function how list all module
'             -
'---------------------------------------------------------------------------------------
'
Public Function vtkExportAll(sourceworkbookname As String)
Dim i As Integer
Dim ttt As String
Dim a As String

a = vtkListAllModules()
i = 0

    While ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + i) <> ""
        a = vtkExportModule(Range(vtkModuleNameRange & vtkFirstLine + i), vtkFirstLine + i, sourceworkbookname)
        i = i + 1
    Wend
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkImportModule
' Author    : user
' Date      : 17/05/2013
' Purpose   : - import module to a workbook
'             - Return number of imported modules
'---------------------------------------------------------------------------------------
'
Public Function vtkImportTestConfig() As Integer
Dim i As Integer

    
        i = 0
    While ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + i) <> ""

        On Error Resume Next
             ' if the module is a class or module
             If (ActiveWorkbook.VBProject.VBComponents(ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + i)).Type = 1 Or ActiveWorkbook.VBProject.VBComponents(ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + i)).Type = 2) Then
                'if the module exist we will delete it and we will replace it
                ActiveWorkbook.VBProject.VBComponents.Remove ActiveWorkbook.VBProject.VBComponents(ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + i))
                ActiveWorkbook.VBProject.VBComponents.Import ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & vtkFirstLine + i)
                ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleInformationsRange & vtkFirstLine + i) = "module imported at " & Now
             
             End If

        i = i + 1
    Wend
vtkImportTestConfig = i
On Error GoTo 0
End Function
