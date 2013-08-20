Attribute VB_Name = "vtkImportExportUtilities"
'---------------------------------------------------------------------------------------
' Module    : vtkImportExportUtilities
' Author    : Jean-Pierre Imbert
' Date      : 07/08/2013
' Purpose   : Group utilitiy functions for Modules Import/Export management
'---------------------------------------------------------------------------------------

Option Explicit

Private vbaUnitModules As Collection

'---------------------------------------------------------------------------------------
' Function  : VBComponentTypeAsString
' Author    : Jean-Pierre Imbert
' Date      : 07/08/2013
' Purpose   : Convert a VBComponent type into a readable string
'             - ActiveX for ActiveX component
'             - Class for VBA class module component
'             - Document for VBA code part of worksheets and workbooks (ThisWorkbook)
'             - Form for User Forms
'             - Standard for a standard (not a class) VBA code module
'---------------------------------------------------------------------------------------
'
Private Function VBComponentTypeAsString(ctype As Integer)
    Select Case ctype
        Case vbext_ct_ActiveXDesigner
            VBComponentTypeAsString = "ActiveX"
        Case vbext_ct_ClassModule
            VBComponentTypeAsString = "Class"
        Case vbext_ct_Document
            VBComponentTypeAsString = "Document"
        Case vbext_ct_MSForm
            VBComponentTypeAsString = "Form"
        Case vbext_ct_StdModule
            VBComponentTypeAsString = "Standard"
        Case Else
            VBComponentTypeAsString = "Unknown"
    End Select
End Function

'---------------------------------------------------------------------------------------
' Function  : extensionForVBComponentType
' Author    : Jean-Pierre Imbert
' Date      : 07/08/2013
' Purpose   : Convert a VBComponent type into a file extension for export
'             - ".cls" for VBA class module and VBA Document modules
'             - ".frm" for User Forms
'             - ".bas" for a standard (not a class) VBA code module
'             - ".???" for ActiveX or Unknown type modules
'---------------------------------------------------------------------------------------
'
Private Function extensionForVBComponentType(ctype As Integer)
    Select Case ctype
        Case vbext_ct_ActiveXDesigner
            extensionForVBComponentType = ".???"
        Case vbext_ct_ClassModule
            extensionForVBComponentType = ".cls"
        Case vbext_ct_Document
            extensionForVBComponentType = ".cls"
        Case vbext_ct_MSForm
            extensionForVBComponentType = ".frm"
        Case vbext_ct_StdModule
            extensionForVBComponentType = ".bas"
        Case Else
            extensionForVBComponentType = ".???"
    End Select
End Function

'---------------------------------------------------------------------------------------
' Function  : vtkStandardCategoryForModuleName
' Author    : Jean-Pierre Imbert
' Date      : 08/08/2013
' Purpose   : return the standard category for a module depending on its name
'             - "VBAUnit" if the module belongs to the VBAUnit list
'             - "Test" if the module name ends with "Tester"
'             - "Prod" if none of the above
'             The standard category is a proposed one, it's not mandatory
'---------------------------------------------------------------------------------------
'
Public Function vtkStandardCategoryForModuleName(moduleName As String) As String
   
   On Error Resume Next
    Dim ret As String
    ret = vtkVBAUnitModulesList.Item(moduleName)
    If err.Number = 0 Then
        vtkStandardCategoryForModuleName = "VBAUnit"
       On Error GoTo 0
        Exit Function
        End If
   On Error GoTo 0
   
    If Right(moduleName, 6) Like "Tester" Then
        vtkStandardCategoryForModuleName = "Test"
       Else
        vtkStandardCategoryForModuleName = "Prod"
    End If

End Function

'---------------------------------------------------------------------------------------
' Function  : vtkStandardPathForModule
' Author    : Jean-Pierre Imbert
' Date      : 08/08/2013
' Purpose   : return the standard relative path to export a module given as a VBComponent
'---------------------------------------------------------------------------------------
'
Public Function vtkStandardPathForModule(module As VBComponent) As String

    Dim path As String
    Select Case vtkStandardCategoryForModuleName(moduleName:=module.name)
        Case "VBAUnit"
            path = "Source\VbaUnit\"
        Case "Prod"
            path = "Source\ConfProd\"
        Case "Test"
            path = "Source\ConfTest\"
    End Select
    
    vtkStandardPathForModule = path & module.name & extensionForVBComponentType(ctype:=module.Type)
    
End Function

'---------------------------------------------------------------------------------------
' Function  : vtkVBAUnitModulesList
' Author    : Jean-Pierre Imbert
' Date      : 07/08/2013
' Purpose   : return a collection initialized with the list of the VBAUnit Modules
'---------------------------------------------------------------------------------------
'
Public Function vtkVBAUnitModulesList() As Collection
    If vbaUnitModules Is Nothing Then
        Set vbaUnitModules = New Collection
        With vbaUnitModules
            .Add Item:="VbaUnitMain", Key:="VbaUnitMain"
            .Add Item:="Assert", Key:="Assert"
            .Add Item:="AutoGen", Key:="AutoGen"
            .Add Item:="IAssert", Key:="IAssert"
            .Add Item:="IResultUser", Key:="IResultUser"
            .Add Item:="IRunManager", Key:="IRunManager"
            .Add Item:="ITest", Key:="ITest"
            .Add Item:="ITestCase", Key:="ITestCase"
            .Add Item:="ITestManager", Key:="ITestManager"
            .Add Item:="RunManager", Key:="RunManager"
            .Add Item:="TestCaseManager", Key:="TestCaseManager"
            .Add Item:="TestClassLister", Key:="TestClassLister"
            .Add Item:="TesterTemplate", Key:="TesterTemplate"
            .Add Item:="TestFailure", Key:="TestFailure"
            .Add Item:="TestResult", Key:="TestResult"
            .Add Item:="TestRunner", Key:="TestRunner"
            .Add Item:="TestSuite", Key:="TestSuite"
            .Add Item:="TestSuiteManager", Key:="TestSuiteManager"
        End With
    End If
    Set vtkVBAUnitModulesList = vbaUnitModules
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkImportOneModule
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Import a module from a file into a project
' Parameters :
'           - project, a VBProject into which import the module
'           - moduleName, the name of module to import
'           - filePath, path of the file to import as new module
'             if the import succeed, the imported module replace the old one if any
'
' Programming Tip
'       It's impossible to remove an existing module by VBA code, so the import is done
'       by a/ really import the module if it doesn't exist
'          b/ erase then rewrite the lines of code if it already exists
'
' WARNING : The code of User Form can be imported with this method
'           but the form layout can't be imported if it already exist in the project
'---------------------------------------------------------------------------------------
'
Public Sub vtkImportOneModule(project As VBProject, moduleName As String, filePath As String)
    Dim newModule As VBComponent, oldModule As VBComponent
    
   
   On Error Resume Next
    Set oldModule = project.VBComponents(moduleName)
    
    ' If the oldModule doesn't exist, we can directly import the file
    If oldModule Is Nothing Then
        Set newModule = project.VBComponents.Import(filePath)
        If Not newModule Is Nothing Then newModule.name = moduleName
       Else
    ' If the oldModule exists, we have to copy lines of code from the file
        ' Read File
        Dim fso, buf As TextStream, code As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        Const ForReading = 1, ForWriting = 2, ForAppending = 3
        Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

        Set buf = fso.OpenTextFile(filePath, ForReading, False, TristateUseDefault)
        
        ' If file open OK, read the code
        If Not buf Is Nothing Then
            ' Discards the first lines before the first "Attribute" line
            Do
                code = buf.ReadLine
                Loop While Not Left$(code, 9) Like "Attribute"
            ' Discards the "Attribute" lines
            Do While Left$(code, 9) Like "Attribute"
                code = buf.ReadLine
                Loop
            ' Read all remaining lines
            code = code & vbCrLf & buf.ReadAll
            
            ' Replace the code
            oldModule.CodeModule.DeleteLines StartLine:=1, Count:=oldModule.CodeModule.CountOfLines
            oldModule.CodeModule.InsertLines 1, code
        End If
        Set fso = Nothing
    End If
   On Error GoTo 0
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkRecreateConfiguration
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : recreate a complete configuration based on
'             - the vtkConfiguration sheet of the project
'             - the exported modules located in the Source folder
' Params    - projectName
'           - configurationName
' WARNING : We use vtkImportOneModule because the document module importation is
'           not efficient with VBComponents.Import (creation of a double class module
'           instead of import the Document code)
' TEST :
'   execute 'vtkRecreateConfiguration projectName :="VBAToolKit",configurationName:="VBAToolKit"'
'   if the Project to reconfigure is installed as AddIn, it must be uninstalled then reinstalled.
'
' IMPORTANT TO DO :
'   il faut éventuellement refermer le classeur qui reste ouvert à la fin (SaveAs n'est pas Close)
'   il faut aussi récupérer les propriétés du projet existant (description dans AddIn)
'   il faut réactiver les références éventuelles (notamment extensions)
'   il faut récupérer la feuille vtkConfigurations existante pour le projet de DEV
'   il faudrait aussi récupérer les modules non exportés du projet DEV (tmptest)
'   il faut tester avec un miniprojet où les deux fichiers Excel sont dans Tests
'---------------------------------------------------------------------------------------
'
Public Sub vtkRecreateConfiguration(projectName As String, configurationName As String)
    Dim cm As vtkConfigurationManager, rootPath As String
    Dim cn As Integer, filePath As String, wbPath As String, wb As Workbook, i As Integer
    ' Get the project and the rootPath of the project
    Set cm = vtkConfigurationManagerForProject(projectName)
    rootPath = cm.rootPath
    ' Get the configuration number in the project and the path of the file
    cn = cm.getConfigurationNumber(configuration:=configurationName)
    wbPath = cm.getConfigurationPathWithNumber(cn)
    ' Create a new Excel file
    Set wb = vtkCreateExcelWorkbook()
    ' Set the projectName
    wb.VBProject.name = configurationName
    ' Import all modules for this configuration from the source directory
    For i = 1 To cm.moduleCount
        filePath = cm.getModulePathWithNumber(numModule:=i, numConfiguration:=cn)
        If Not filePath Like "" Then vtkImportOneModule project:=wb.VBProject, moduleName:=cm.module(i), filePath:=rootPath & "\" & filePath
    Next i
    ' Save the Excel file with the good type and erase the previous one
    wb.SaveAs Filename:=rootPath & "\" & wbPath, FileFormat:=vtkDefaultFileFormat(wbPath)
End Sub

'
''---------------------------------------------------------------------------------------
'' Procedure : vtkListAllModules
'' Author    : user
'' Date      : 17/05/2013
'' Purpose   : - call VtkInitializeExcelfileWithVbaUnitModuleName and use his return value
''             - list all module of current project , verify that the module
''              is not a vbaunit and write his name in the range
''
''---------------------------------------------------------------------------------------
''
'Public Function vtkListAllModules() As Integer
'Dim i As Integer
'Dim j As Integer
'Dim k As Integer
'Dim t As Integer
'
't = VtkInitializeExcelfileWithVbaUnitModuleName()
'k = 0
'  For i = 1 To ActiveWorkbook.VBProject.VBComponents.Count
'    If vtkIsVbaUnit(ActiveWorkbook.VBProject.VBComponents.Item(i).name) = False Then
'        ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & t + k) = ActiveWorkbook.VBProject.VBComponents.Item(i).name
'        ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & t + k).Interior.ColorIndex = 8
'        k = k + 1
'    End If
'  Next
'vtkListAllModules = k
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : vtkCreateModuleFile
'' Author    : user
'' Date      : 17/05/2013
'' Purpose   : - this function allow to create a file
''             - return message contain informations: time , file created or replaced
''---------------------------------------------------------------------------------------
''
'Public Function vtkCreateModuleFile(fullPath As String) As String
'
'Dim fso As New FileSystemObject
'
'If fso.FileExists(fullPath) = False Then
'    fso.CreateTextFile (fullPath)
'vtkCreateModuleFile = "File created successfully at" & Now
'Else
'vtkCreateModuleFile = "File last update at" & Now
'End If
'End Function
'
'
'
''---------------------------------------------------------------------------------------
'' Procedure : vtkExportModule
'' Author    : user
'' Date      : 14/05/2013
'' Purpose   : - function take modulename , and line number , and workbookSource Name
''             - create files of modules if they don't exist ,or update it
''             - export module to the right folders  (documents , worksheets)
''             - write creation file informations
''             - write exported file location
''
''  if "vbaunitclass" then
''       if vbaUnitMain then ===================>path= vbaunit ".bas"
''       else                ===================>path= vbaunit ".cls"
''       endif
''  else
''     case module.type
''
''       1.module ,to ===========================>path= confprod ".BAS"
''       2.classmodule, if---nameTester to ======>path= ConfTest ".CLS"
''                      else ====================>path= ConfProd ".CLS"
''       3.Form   ,to ===========================>path= confprod ".FRM"
''     sheet ,worksheet, workbook ===============> do nothing
''  endif
''  vtkCreateModuleFile(path)
''  sheet.range = path
''---------------------------------------------------------------------------------------
''
'Public Function vtkExportModule(modulename As String, lineNumber As Integer, sourceworkbook As String) As String
'
' Dim fullPath As String
' Dim path As String
' Dim MsgCreationFile As String
' Dim Test As String
' Dim DevPath As String
' Dim DelivPath As String
' Dim color As Integer
' color = 2
' Dim fso As New FileSystemObject
' path = fso.GetParentFolderName(ActiveWorkbook.path)
'
'
'    If vtkIsVbaUnit(modulename) = True Then
'          If modulename = "VbaUnitMain" Then
'                fullPath = path & "\Source\VbaUnit\" & modulename & ".bas"  'full path of file that will be created
'                DevPath = fullPath
'                DelivPath = ""
'                color = 3
'          Else
'                fullPath = path & "\Source\VbaUnit\" & modulename & ".cls"  'full path of file that will be created
'                DevPath = fullPath
'                DelivPath = ""
'                color = 3
'          End If
'    Else
'
'        On Error Resume Next
'
'
'    Select Case Workbooks(sourceworkbook).VBProject.VBComponents(modulename).Type
'
'        Case 1 '1module : export to confprod
'
'           fullPath = path & "\Source\ConfProd\" & modulename & ".bas"  'full path of file that will be created
'            DevPath = fullPath
'            DelivPath = fullPath
'
'
'        Case 2 '2 class module : export to ConfTest or ConfProd
'
'            If Right(modulename, 6) Like "Tester" Then ' verify if modulename end is like Tester
'
'                ' This Document is a test module export to confTest
'                fullPath = path & "\Source\ConfTest\" & modulename & ".CLS"
'                DevPath = fullPath
'                DelivPath = ""
'                color = 3
'            Else
'
'                'the document is a classmodule export to confprod
'                fullPath = path & "\Source\ConfProd\" & modulename & ".CLS"
'                DevPath = fullPath
'                DelivPath = fullPath
'            End If
'        Case 3 '3 forms
'
'                'the document is a classmodule export to confprod
'                fullPath = path & "\Source\ConfProd\" & modulename & ".FRM"
'                DevPath = fullPath
'                DelivPath = fullPath
'
'        Case 100 'excel sheets , we will not export them for the moment
'                DevPath = ""
'                DelivPath = ""
'                color = 3
'                ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & lineNumber).Interior.ColorIndex = color
'
'        Case Else 'normally we haven't other type but if we find another type we will export it to main project folder
'                DevPath = ""
'                DelivPath = ""
'                color = 3
'                ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & lineNumber).Interior.ColorIndex = color
'          Exit Function
'
'      End Select
'    End If
'
'   MsgCreationFile = vtkCreateModuleFile(fullPath)
'   Workbooks(sourceworkbook).VBProject.VBComponents(modulename).Export (fullPath) 'export module to the right folder
'
'   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & lineNumber) = DevPath
'
'   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & lineNumber) = DelivPath
'   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & lineNumber).Interior.ColorIndex = color
'
'   ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkInformationRange & lineNumber) = MsgCreationFile
'
'   On Error GoTo 0
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : vtkExportAll
'' Author    : user
'' Date      : 16/05/2013
'' Purpose   : - call function how list all module
''             -
''---------------------------------------------------------------------------------------
''
'Public Function vtkExportAll(sourceworkbookname As String)
'Dim i As Integer
'Dim ttt As String
'Dim a As String
'
'a = vtkListAllModules()
'i = 0
'
'    While ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine + i) <> ""
'        a = vtkExportModule(Range(vtkModuleNameRange & vtkFirstLine + i), vtkFirstLine + i, sourceworkbookname)
'        i = i + 1
'    Wend
'End Function
'
