Attribute VB_Name = "vtkImportExportUtilities"
'---------------------------------------------------------------------------------------
' Module    : vtkImportExportUtilities
' Author    : Jean-Pierre Imbert
' Date      : 07/08/2013
' Purpose   : Group utilitiy functions for Modules Import/Export management
'
' Copyright 2013 Skwal-Soft (http://skwalsoft.com)
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
'---------------------------------------------------------------------------------------

Option Explicit

Private vbaUnitModules As Collection
Public Const VTK_UNKNOWN_MODULE = 4000

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
    If Err.Number = 0 Then
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
'           - project, a VBProject into which to import the module
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
' Procedure : vtkExportOneModule
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Export a module from a project to a file
' Parameters :
'           - project, a VBProject from which to export the module
'           - moduleName, the name of module to export
'           - filePath, path of the file to export the module
'           - normalize, if True (default value) the token are nomalized
' NOTE      : The file in which to export is deleted prior to export if it already exists
'             except if the file has readonly attribute
' NOTE      : The tokens in the exported module are normalized (default behavior)
'---------------------------------------------------------------------------------------
'
Public Sub vtkExportOneModule(project As VBProject, moduleName As String, filePath As String, Optional normalize As Boolean = True)
    Dim fso As New FileSystemObject, m As VBComponent
    
   On Error GoTo vtkExportOneModule_Error
   
    ' Get the module to export
    Set m = project.VBComponents(moduleName)
        
    ' Kill file if it already exists only AFTER get the module, if it not exists the file must not be deleted
    If fso.FileExists(filePath) Then fso.DeleteFile fileSpec:=filePath
    
    ' Create full path if needed
    vtkCreateFolderPath fileOrFolderPath:=filePath
    
    ' Export module
    m.Export fileName:=filePath
    
    ' Normalize tokens in file
    If normalize Then vtkNormalizeFile filePath, vtkListOfProperlyCasedIdentifiers
    
   On Error GoTo 0
    Exit Sub

vtkExportOneModule_Error:
    If Err.Number = 9 Then
        Err.Raise Number:=VTK_UNKNOWN_MODULE, Source:="ExportOneModule", Description:="Module to export doesn't exist : " & moduleName
       Else
        Err.Raise Number:=VTK_UNEXPECTED_ERROR, Source:="ExportOneModule", Description:="Unexpected error when exporting " & moduleName & " : " & Err.Description
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkExportModulesFromAnotherProject
' Author    : Jean-Pierre Imbert
' Date      : 22/08/2013
' Purpose   : Export modules listed in a configuration for a project
'             - the project/configuration containing the modules list are projectName/confName parameters
'             - the modules are extracted from project projectWithModules
'             - the modules are exported to pathes listed in project/configuration
' NOTE      : Used to get VBAUnit modules (and libs ?) when creating a project
'---------------------------------------------------------------------------------------
'
Public Sub vtkExportModulesFromAnotherProject(projectWithModules As VBProject, projectName As String, confName As String)
    Dim cm As vtkConfigurationManager, rootPath As String
    Dim cn As Integer, filePath As String, i As Integer
    
   On Error GoTo vtkExportModulesFromAnotherProject_Error

    ' Get the project and the rootPath of the project
    Set cm = vtkConfigurationManagerForProject(projectName)
    cn = cm.getConfigurationNumber(configuration:=confName)
    rootPath = cm.rootPath
    
    ' Export all modules for this configuration from the projectWithModules
    For i = 1 To cm.moduleCount
        filePath = cm.getModulePathWithNumber(numModule:=i, numConfiguration:=cn)
        If Not filePath Like "" Then vtkExportOneModule project:=projectWithModules, moduleName:=cm.module(i), filePath:=rootPath & "\" & filePath
    Next i
    
   On Error GoTo 0
   Exit Sub

vtkExportModulesFromAnotherProject_Error:
    Err.Raise VTK_UNEXPECTED_ERROR, "vtkExportModulesFromAnotherProject", "Unexpected error when exporting modules from " & projectWithModules.name & " : " & Err.Description
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkImportModulesInAnotherProject
' Author    : Jean-Pierre Imbert
' Date      : 22/08/2013
' Purpose   : Import modules listed in a configuration for a project
'             - the project/configuration containing the modules list are projectName/confName parameters
'             - the modules are imported in the project projectForModules
'             - the modules are imported from pathes listed in project/configuration
'---------------------------------------------------------------------------------------
'
Public Sub vtkImportModulesInAnotherProject(projectForModules As VBProject, projectName As String, confName As String, Optional cm As vtkConfigurationManager = Nothing)
    Dim rootPath As String, cn As Integer, filePath As String, i As Integer
    
   On Error GoTo vtkImportModulesInAnotherProject_Error

    ' Get the project and the rootPath of the project
    If cm Is Nothing Then Set cm = vtkConfigurationManagerForProject(projectName)
    cn = cm.getConfigurationNumber(configuration:=confName)
    rootPath = cm.rootPath
    
    ' Import all modules for this configuration into the projectForModules
    For i = 1 To cm.moduleCount
        filePath = cm.getModulePathWithNumber(numModule:=i, numConfiguration:=cn)
        If Not filePath Like "" Then vtkImportOneModule project:=projectForModules, moduleName:=cm.module(i), filePath:=rootPath & "\" & filePath
    Next i
    
   On Error GoTo 0
   Exit Sub

vtkImportModulesInAnotherProject_Error:
    Err.Raise VTK_UNEXPECTED_ERROR, "vtkImportModulesInAnotherProject_Error", "Unexpected error when importing modules into " & projectForModules.name & " : " & Err.Description
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkRecreateConfiguration
' Author    : JPI-Conseil
' Purpose   : Recreate a complete configuration based on
'             - the vtkConfiguration sheet of the project
'             - the exported modules located in the Source folder
'
' Params    : - projectName
'             - configurationName
'             - Optional Configuration Manager
'                   if no Conf Manager are given, the standard Conf manager (Excel)
'                   of the project is used
'
' Raises    : - VTK_UNEXPECTED_ERROR
'             - VTK_WORKBOOK_ALREADY_OPEN
'             - VTK_NO_SOURCE_FILES
'             - VTK_WRONG_FILE_PATH
'
' WARNING : We use vtkImportOneModule because the document module importation is
'           not efficient with VBComponents.Import (creation of a double class module
'           instead of import the Document code)
'
'---------------------------------------------------------------------------------------
'
Public Sub vtkRecreateConfiguration(projectName As String, configurationName As String, Optional confManager As vtkConfigurationManager = Nothing)
    Dim cm As vtkConfigurationManager
    Dim rootPath As String
    Dim wbPath As String, templatePath As String
    Dim Wb As Workbook
    Dim tmpWb As Workbook
    Dim fso As New FileSystemObject

    On Error GoTo vtkRecreateConfiguration_Error
    Application.EnableEvents = False
    
    ' Get the Conf Manager and the rootPath of the project
    If confManager Is Nothing Then
        Set cm = vtkConfigurationManagerForProject(projectName)
       Else
        Set cm = confManager
    End If
    rootPath = cm.rootPath
    
    ' Get the actual path of the file (without template if any)
    wbPath = cm.configurations(configurationName).path
    
    ' Make sure the workbook we want to create is not open
    ' NB : open add-ins don't count and are managed further down
    For Each tmpWb In Workbooks
        If tmpWb.name Like fso.GetFileName(wbPath) Then Err.Raise VTK_WORKBOOK_ALREADY_OPEN
    Next
    
    'Make sure the source files exist
    Dim mo As vtkModule
    Dim conf As vtkConfiguration
    Set conf = cm.configurations(configurationName)
    For Each mo In conf.modules
        If fso.FileExists(rootPath & "\" & mo.getPathForConfiguration(configurationName)) = False Then
            Err.Raise VTK_NO_SOURCE_FILES
        End If
    Next
    
    ' Create a new Excel file
    templatePath = conf.template
    If templatePath = "" Then
        Set Wb = vtkCreateExcelWorkbook()   ' If there is no template, a new workbook is created
       Else
        templatePath = rootPath & "\" & templatePath
        If Not fso.FileExists(templatePath) Then
            Err.Raise VTK_TEMPLATE_NOT_FOUND
        End If
        Set Wb = Workbooks.Open(fileName:=templatePath, ReadOnly:=True) ' If there is a template, it's open as ReadOnly
    End If
    
    ' Set the projectName
    Wb.VBProject.name = conf.projectName
    
    ' Set the Workbook properties
    Wb.BuiltinDocumentProperties("Title").Value = conf.projectName
    Wb.BuiltinDocumentProperties("Comments").Value = conf.comment
    
    ' Import all modules for this configuration from the source directory
    vtkImportModulesInAnotherProject projectForModules:=Wb.VBProject, projectName:=projectName, confName:=configurationName, cm:=cm
    
    ' Recreate references in the new Excel File
    conf.addReferencesToWorkbook Wb
    
    ' Duplicate Conf Manager if DEV configuration
    If conf.isDEV Then
        Dim cmE As New vtkConfigurationManagerExcel
        cmE.duplicate Wb, cm
    End If
    
    ' Protect VBA Project with optional password
    If conf.password <> "" Then vtkProtectProject project:=Wb.VBProject, password:=conf.password
    
    ' VB will not let the workbook be saved under the name of an already opened workbook, which
    ' is annoying when recreating an add-in (always opened). The following code works around this.
    Dim tmpPath As String
    ' Add a random string to the file name of the workbook that will be saved
    tmpPath = fso.BuildPath(rootPath & "\" & fso.GetParentFolderName(wbPath), _
              vtkStripFilePathOrNameOfExtension(fso.GetFileName(wbPath)) & _
              CStr(Round((99999 - 10000 + 1) * Rnd(), 0)) + 10000 & _
              "." & fso.GetExtensionName(wbPath))
    
    ' Create the folder containing the workbook if a 1-level or less deep folder structure
    ' is specified in the configuration path.
    vtkCreateFolderPath tmpPath
    
    ' Save the new workbook with the correct extension
    Wb.IsAddin = vtkDefaultIsAddIn(wbPath)
    Wb.SaveAs fileName:=tmpPath, FileFormat:=vtkDefaultFileFormat(wbPath)
    Wb.Close saveChanges:=False
    Application.EnableEvents = True
    
    ' Delete the old workbook if it exists
    Dim fullWbpath As String
    fullWbpath = rootPath & "\" & wbPath
    If fso.FileExists(fullWbpath) Then fso.DeleteFile (fullWbpath)
    
    ' Rename the new workbook without the random string
    fso.GetFile(tmpPath).name = fso.GetFileName(rootPath & "\" & wbPath)
    
    On Error GoTo 0
    Exit Sub

vtkRecreateConfiguration_Error:

    If Not Wb Is Nothing Then Wb.Close saveChanges:=False

    Err.Source = "vtkRecreateConfiguration of module vtkImportExportUtilities"
    
    Select Case Err.Number
        Case VTK_WORKBOOK_ALREADY_OPEN
            Err.Number = VTK_WORKBOOK_ALREADY_OPEN
            Err.Description = "The configuration you're trying to create (" & configurationName & ") corresponds to an open workbook. " & _
                              "Please close it before recreating the configuration."
        Case VTK_TEMPLATE_NOT_FOUND
            Err.Number = VTK_TEMPLATE_NOT_FOUND
            Err.Description = "The configuration you're trying to create (" & configurationName & ") needs an Excel template file (" & templatePath & "). " & _
                              "This template file is unreachable."
        Case VTK_NO_SOURCE_FILES
            Err.Number = VTK_NO_SOURCE_FILES
            Err.Description = "The configuration you're trying to create (" & configurationName & ") is missing one or several source files." & _
                              "Please export the modules in their relevant path before recreating the configuration."
        Case VTK_WRONG_FILE_PATH
            Err.Number = VTK_WRONG_FILE_PATH
            Err.Description = "The configuration you're trying to create (" & configurationName & ") has a invalid path." & _
                              "Please check if the folder structure it needs is not more than one-level deep."
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select

    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkExportConfiguration
' Author    : Lucas Vitorino
' Purpose   : - Export all the modules of a configuration to their respective paths, and returns the number
'               of exported modules.
'             - Optionally, export only the modules modified since last save
' Behaviour : - If onlyModified = True, export modules whose source file does not yet exist and modules whose file
'               exists but that have been modified since last save.
'             - If onlyModified = False, export every module of the configuration.
' Returns   : Integer
'---------------------------------------------------------------------------------------
'
Public Function vtkExportConfiguration(projectWithModules As VBProject, projectName As String, confName As String, _
        Optional onlyModified As Boolean = False) As Integer
    
    Dim cm As vtkConfigurationManager, rootPath As String
    Dim mo As vtkModule
    Dim exportedModulesCount As Integer: exportedModulesCount = 0
    Dim fso As New FileSystemObject
    
    On Error GoTo vtkExportConfiguration_Error

    ' Export all modules for this configuration from the projectWithModules
    Set cm = vtkConfigurationManagerForProject(projectName)
    
    For Each mo In cm.configurations(confName).modules
        
        Dim modulePath As String
        modulePath = cm.rootPath & "\" & mo.path
                   
        ' The conditon to export could be simplified in
        '   If Not (onlyModified And fso.FileExists(modulePath) And mo.VBAModule) Then <export>
        ' but it doesn't seem to work, so we stick to this code.
        If onlyModified And fso.FileExists(modulePath) Then
            If mo.VBAModule.Saved = False Then
                vtkExportOneModule projectWithModules, mo.name, modulePath
                exportedModulesCount = exportedModulesCount + 1
            End If
        Else
            vtkExportOneModule projectWithModules, mo.name, modulePath
            exportedModulesCount = exportedModulesCount + 1
        End If
        
    Next
    
    On Error GoTo 0
    vtkExportConfiguration = exportedModulesCount
    Exit Function

vtkExportConfiguration_Error:
    Err.Raise Err.Number, "procedure vtkExportConfiguration of Module vtkImportExportUtilities", Err.Description
    Resume Next

End Function
