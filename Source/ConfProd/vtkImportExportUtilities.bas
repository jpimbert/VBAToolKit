Attribute VB_Name = "vtkImportExportUtilities"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : vtkConfSheet
' Author    : Abdelfattah Lahbib
' Date      : 14/05/2013
' Purpose   : - Create new sheet (if it does not already exist) that will contain
'               the table of parameters
'---------------------------------------------------------------------------------------
'
Public Function vtkConfSheet() As String
    
    Dim sheetName
    sheetName = "configurations"
    
    On Error Resume Next
        Worksheets(sheetName).Select
    If Err <> 0 Then
        Worksheets.Add.name = sheetName
    End If
    
    vtkConfSheet = sheetName
    On Error GoTo 0

End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkModuleNameRange
' Author    : Abdelfattah Lahbib
' Date      : 13/05/2013
' Purpose   : - Return name of the range (column of the vtkConfigurations worksheet in the Project_DEV
'               Excel file) that contains the names of the modules
'             - Write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleNameRange() As String
    vtkModuleNameRange = "A"
    ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine - 2) = "Module Name"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkModuleDevRange
' Author    : Abdelfattah Lahbib
' Date      : 13/05/2013
' Purpose   : - Return name of the range (column of the vtkConfigurations worksheet in the Project_DEV
'               Excel file) that contains the list of the paths of the developement configuration
'             - Write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleDevRange() As String
    vtkModuleDevRange = "B"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkModuleDeliveryRange
' Author    : Abdelfattah Lahbib
' Date      : 13/05/2013
' Purpose   : - Return name of the range (column of the vtkConfigurations worksheet in the Project_DEV
'               Excel file) that contains the list of the paths of the delivery configuration
'             - Write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleDeliveryRange() As String
    vtkModuleDeliveryRange = "C"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInformationRange
' Author    : Abdelfattah Lahbib
' Date      : 13/05/2013
' Purpose   : - Return name of the range (column of the vtkConfigurations worksheet in the Project_DEV
'               Excel file) that contains information
'             - Write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkInformationRange() As String
    vtkInformationRange = "D"
    ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkInformationRange & vtkFirstLine - 3) = "File Informations"
End Function
'---------------------------------------------------------------------------------------
' Procedure : vtkModuleInformationsRange
' Author    : Abdelfattah Lahbib
' Date      : 13/05/2013
' Purpose   : - Return name of the range (column of the vtkConfigurations worksheet in the Project_DEV
'               Excel file) that contains information about the modules
'             - Write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleInformationsRange() As String
    vtkModuleInformationsRange = "E"
    ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleInformationsRange & vtkFirstLine - 3) = "Modules Informations"
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkFirstLine
' Author    : Abdelfattah Lahbib
' Date      : 13/05/2013
' Purpose   : - Define the start line
'---------------------------------------------------------------------------------------
'
Public Function vtkFirstLine() As Integer
    vtkFirstLine = 4
End Function
'---------------------------------------------------------------------------------------
' Procedure : VtkInitializeExcelfileWithVbaUnitModuleName
' Author    : Abdelfattha Lahbib
' Date      : 09/05/2013
' Purpose   : - Initialize the vtkConfigurations worksheet with the modules of VBA Unit
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
' Author    : Abdelfattah Lahbib
' Date      : 17/05/2013
' Purpose   : - Take module name in parameter and check if it belongs to VBA Unit
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
' Author    : Abdelfattah Lahbib
' Date      : 17/05/2013
' Purpose   : - Call VtkInitializeExcelfileWithVbaUnitModuleName and use his return value
'             - List all modules of the current project, check if they don't belong to VBA Unit,
'               and write their name in the range.
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
' Author    : Abdelfattah Lahbib
' Date      : 17/05/2013
' Purpose   : - Create a module file
'             - Return message that contains information : time of the operation, file created
'               or updated.
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateModuleFile(fullPath As String) As String

    Dim fso As New FileSystemObject
    
    If fso.FileExists(fullPath) = False Then
        fso.CreateTextFile (fullPath)
        vtkCreateModuleFile = "File created successfully at" & Now
    Else
        vtkCreateModuleFile = "File last update at" & Now
    End If
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : vtkExportModule
' Author    : Abdelfattah Lahbib
' Date      : 14/05/2013
' Purpose   : - Parameters : module name, line number, and workbookSource name
'             - Create module file if it doesn't exist, or update it
'             - Export module file to the right folders  (documents , worksheets)
'             - Write information about the file operation (creation, update, time)
'             - Write exported file location
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

 Dim fullPath As String
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
                fullPath = path & "\Source\VbaUnit\" & modulename & ".bas"  'full path of file that will be created
                DevPath = fullPath
                DelivPath = ""
                color = 3
          Else
                fullPath = path & "\Source\VbaUnit\" & modulename & ".cls"  'full path of file that will be created
                DevPath = fullPath
                DelivPath = ""
                color = 3
          End If
    Else
        
        On Error Resume Next
    
    
    Select Case Workbooks(sourceworkbook).VBProject.VBComponents(modulename).Type
          
        Case 1 '1 module : export to confprod
         
           fullPath = path & "\Source\ConfProd\" & modulename & ".bas"  'full path of file that will be created
            DevPath = fullPath
            DelivPath = fullPath
          
         
        Case 2 '2 class module : export to ConfTest or ConfProd
         
            If Right(modulename, 6) Like "Tester" Then ' verify if modulename end is like Tester
                    
                ' This Document is a test module export to confTest
                fullPath = path & "\Source\ConfTest\" & modulename & ".CLS"
                DevPath = fullPath
                DelivPath = ""
                color = 3
            Else
    
                'the document is a classmodule export to confprod
                fullPath = path & "\Source\ConfProd\" & modulename & ".CLS"
                DevPath = fullPath
                DelivPath = fullPath
            End If
        Case 3 '3 forms
        
                'the document is a classmodule export to confprod
                fullPath = path & "\Source\ConfProd\" & modulename & ".FRM"
                DevPath = fullPath
                DelivPath = fullPath
                
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

   MsgCreationFile = vtkCreateModuleFile(fullPath)
   Workbooks(sourceworkbook).VBProject.VBComponents(modulename).Export (fullPath) 'export module to the right folder

   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & lineNumber) = DevPath
   
   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & lineNumber) = DelivPath
   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & lineNumber).Interior.ColorIndex = color
   
   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkInformationRange & lineNumber) = MsgCreationFile
 
   On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkExportAll
' Author    : Abdelfattah Lahbib
' Date      : 16/05/2013
' Purpose   : - Call function to list all modules
'             - Call function to export each module
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
' Author    : Abdelfattah Lahbib
' Date      : 17/05/2013
' Purpose   : - Import module to a workbook
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
                'if the module exists, delete it and replace it
                ActiveWorkbook.VBProject.VBComponents.Remove ActiveWorkbook.VBProject.VBComponents(ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + i))
                ActiveWorkbook.VBProject.VBComponents.Import ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & vtkFirstLine + i)
                ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleInformationsRange & vtkFirstLine + i) = "module imported at " & Now
             
             End If

        i = i + 1
    Wend

    vtkImportTestConfig = i
    On Error GoTo 0
    
End Function
