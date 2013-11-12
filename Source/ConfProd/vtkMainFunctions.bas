Attribute VB_Name = "vtkMainFunctions"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkMainFunctions
' Author    : Jean-Pierre Imbert
' Date      : 04/07/2013
' Purpose   : This module contains the functions called for the main capacities of VBAToolKit
'               - new project creation
'               - (other capacities will be delelopped later)
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

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateProject
' Author    : JPI-Conseil
' Date      : 03/04/2013
' Purpose   : Create a new project managed with VBAToolKit
'               - create the tree folder for a new project
'                   - Source containing ConfProd, ConfTest and VBAUnit
'                   - Project containing the main Excel file for the project
'                   - an empty Tests folder
'               - Initialize the Git repository for the project
'               - Create Xlsm Dev and Delivery workbooks
'               - Activate needed VB References
'                   - in the DEV workbook, a reference to the VBAToolKit workbook creating the project is also activated.
'                     During normal use of VBAToolKit, the reference is made to the add-in.
'                     During tests while developing VBAToolKit, the reference is made to the VBAToolKit_DEV workbook.
'                   - the Delivery workbook does not need a reference to VBAToolKit.
'
' Parameters :
'             - path, string containing the path of folder in which to create the project
'             - name, string containing the name of the project to create
'             - displayError, boolean true if dialog box for errors are have to be displayed
'               (used for automatic test where error displaying is disabled)
' Return    : Long error number
' Warning   : - The VB reference for "Microsoft Visual Basic For Application Extensibility" must be activated
'             - unpredictable behavior when creating a new project whose name is used by an existing project
'
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateProject(path As String, name As String, Optional displayError As Boolean = True) As Long
    
  On Error GoTo vtkCreateProject_Error

    Dim fso As New FileSystemObject

    ' Create the vtkProject object attached to the new project
    Dim project As vtkProject
    Set project = vtkProjectForName(projectName:=name)
    Dim rootPath As String
    rootPath = fso.BuildPath(path, project.projectName)
    
    ' Create tree folder
    Dim internalError As Long
    internalError = vtkCreateTreeFolder(rootPath)
    If internalError <> VTK_OK Then GoTo vtkCreateProject_ErrorTreeFolder
    
    ' Create the XML vtkConfigurations sheet in the standard folder, the dtd is supposed to be in the same folder
    createInitializedXMLSheetForProject sheetPath:=fso.BuildPath(rootPath, project.XMLConfigurationStandardRelativePath), _
                                        projectName:=project.projectName, _
                                        dtdPath:="vtkConfigurationsDTD.dtd"
                                        
    ' Create the DTD sheet for the XML vtkConfigurations sheet in the same folder
    createDTDForVtkConfigurations sheetPath:=fso.BuildPath(rootPath, fso.BuildPath(fso.GetParentFolderName(project.XMLConfigurationStandardRelativePath), "vtkConfigurationsDTD.dtd"))
    
    ' Insert the BeforeSave handler in the newly created project
    ' /!\ For that we need to manage the "ThisWorkbook" object, we will do it later
    ' vtkAddBeforeSaveHandlerInDEVWorkbook Wb:=Wb, projectName:=project.projectName, confName:=project.projectDEVName
        
    ' Add the newly created project to the list of projects remembered by VBAToolKit
    vtkAddRememberedProject projectName:=project.projectName, _
                            rootFolder:=rootPath, _
                            xmlRelativePath:=project.XMLConfigurationStandardRelativePath
    
    
    ' Get VBAUnit modules from VBAToolkit (This workbook = current running code)
    vtkExportModulesFromAnotherProject projectWithModules:=ThisWorkbook.VBProject, projectName:=project.projectName, confName:=project.projectDEVName

    
    ' Initialize git
    On Error GoTo vtkCreateProject_ErrorGit
    vtkInitializeGit rootPath

    On Error GoTo 0
    vtkCreateProject = VTK_OK
    Exit Function

vtkCreateProject_ErrorTreeFolder:
    vtkCreateProject = internalError
    If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
    Exit Function
vtkCreateProject_Error:
    vtkCreateProject = Err.Number
    If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
vtkCreateProject_ErrorGit:
    vtkCreateProject = Err.Number
    If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"

End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkRecreateConfiguration
' Author    : JPI-Conseil
' Purpose   : Recreate a complete configuration based on
'             - the vtkConfiguration sheet of the project
'             - the exported modules located in the Source folder
'
' Params    : - projectName
'             - configurationName
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
Public Sub vtkRecreateConfiguration(projectName As String, configurationName As String)
    Dim cm As vtkConfigurationManager
    Dim rootPath As String
    Dim wbPath As String
    Dim Wb As Workbook
    Dim tmpWb As Workbook
    Dim fso As New FileSystemObject

    On Error GoTo vtkRecreateConfiguration_Error
    
    ' Get the project and the rootPath of the project
    Set cm = vtkConfigurationManagerForProject(projectName)
    rootPath = cm.rootPath
    
    ' Get the configuration number in the project and the path of the file
    wbPath = cm.getConfigurationPath(configuration:=configurationName)
    
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
    
    ' Create a new Excel file andset the name of the VBProject, if there is no template
    If conf.templatePath = "" Then
        Set Wb = vtkCreateExcelWorkbook()
        Wb.VBProject.name = configurationName
    Else
        Set Wb = Workbooks.Open(fso.BuildPath(rootPath, conf.templatePath))
    End If
    
    ' Import all modules for this configuration from the source directory
    vtkImportModulesInAnotherProject projectForModules:=Wb.VBProject, projectName:=projectName, confName:=configurationName
    
    ' Recreate references in the new Excel File
    On Error GoTo vtkRecreateConfiguration_referenceError
    Dim tmpRef As vtkReference
    For Each tmpRef In conf.references
        If tmpRef.guid <> "" Then
            Wb.VBProject.references.AddFromGuid tmpRef.guid, 0, 0
        ElseIf tmpRef.path <> "" Then
            Wb.VBProject.references.AddFromFile tmpRef.path
        End If
    Next

    On Error GoTo vtkRecreateConfiguration_Error
    
    ' Set the title and the comments
    Wb.BuiltinDocumentProperties("Title").Value = conf.title
    Wb.BuiltinDocumentProperties("Comments").Value = conf.comment
    
    ' VB will not let the workbook be saved under the name of an already opened workbook, which
    ' is annoying when recreating an add-in (always opened). The following code works around this.
    Dim tmpPath As String
    ' Add a random string to the file name of the workbook that will be saved
    tmpPath = fso.BuildPath(rootPath & "\" & fso.GetParentFolderName(wbPath), _
              vtkStripFilePathOrNameOfExtension(fso.GetFileName(wbPath)) & _
              CStr(Round((99999 - 10000 + 1) * Rnd(), 0)) + 10000 & _
              "." & fso.GetExtensionName(wbPath))

    ' Create the the folder containing the workbook if a 1-level or less deep folder structure
    ' is specified in the configuration path.
    vtkCreateFolderPath tmpPath
    
    ' Without this line, an xla file is not created with the right format
    If vtkDefaultFileFormat(wbPath) = xlAddIn Then Wb.IsAddin = True
    
    ' Save the new workbook with the correct extension
    Wb.SaveAs fileName:=tmpPath, FileFormat:=vtkDefaultFileFormat(wbPath)
    Wb.Close saveChanges:=False
    
    ' Delete the old workbook if it exists
    Dim fullWbpath As String
    fullWbpath = rootPath & "\" & wbPath
    If fso.FileExists(fullWbpath) Then fso.DeleteFile (fullWbpath)
    
    ' Rename the new workbook without the random string
    fso.GetFile(tmpPath).name = fso.GetFileName(rootPath & "\" & wbPath)
    
    On Error GoTo 0
    Exit Sub

vtkRecreateConfiguration_referenceError:
    Err.Number = VTK_REFERENCE_ERROR
    GoTo vtkRecreateConfiguration_Error

vtkRecreateConfiguration_Error:

    If Not Wb Is Nothing Then Wb.Close saveChanges:=False

    Err.Source = "vtkRecreateConfiguration of module vtkImportExportUtilities"
    
    Select Case Err.Number
        Case VTK_REFERENCE_ERROR
            Err.Number = VTK_REFERENCE_ERROR
            Err.Description = "There was a problem activating reference " & tmpRef.name & ""
        Case VTK_WORKBOOK_ALREADY_OPEN
            Err.Number = VTK_WORKBOOK_ALREADY_OPEN
            Err.Description = "The configuration you're trying to create (" & configurationName & ") corresponds to an open workbook. " & _
                              "Please close it before recreating the configuration."
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
