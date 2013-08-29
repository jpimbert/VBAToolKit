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

    ' Create the vtkProject object attached to the new project
    Dim project As vtkProject
    Set project = vtkProjectForName(projectName:=name)
    Dim rootPath As String
    rootPath = path & "\" & project.projectName
    
    ' Create tree folder
    Dim internalError As Long
    internalError = vtkCreateTreeFolder(rootPath)
    If internalError <> VTK_OK Then GoTo vtkCreateProject_ErrorTreeFolder
     
    'Save created project with xlsm extention
    Workbooks.Add.SaveAs (rootPath & "\" & project.projectDEVStandardRelativePath), FileFormat:=xlOpenXMLWorkbookMacroEnabled
    'Rename Project
    Workbooks(project.workbookDEVName).VBProject.name = project.projectDEVName
    'call function who activate references
    VtkActivateReferences wb:=Workbooks(project.workbookDEVName)
    'initialize configuration Sheet with VBAUnit modules
    vtkInitializeVbaUnitNamesAndPathes project:=project.projectName
    ' Save Development Project Workbook
    Workbooks(project.workbookDEVName).save
    
    'Create delivery workbook
    Workbooks.Add.SaveAs (rootPath & "\" & project.projectStandardRelativePath), FileFormat:=(52) '52 is xlsm format
    'Rename Project
    Workbooks(project.workbookName).VBProject.name = project.projectName
    'call function who activate references
    VtkActivateReferences wb:=Workbooks(project.workbookName)
    ' A module must be added in the Excel File for the project parameters to be saved
    Workbooks(project.workbookName).VBProject.VBComponents.Add ComponentType:=vbext_ct_StdModule
    ' Save and Close Delivery Project WorkBook
    Workbooks(project.workbookName).Close savechanges:=True
    
    Dim wb As Workbook
    Set wb = Workbooks(project.workbookDEVName)
    wb.Activate
    ' Get VBAUnit modules from VBAToolkit (This workbook = current running code)
    vtkExportModulesFromAnotherProject projectWithModules:=ThisWorkbook.VBProject, projectName:=project.projectName, confName:=project.projectDEVName
    ' Import VBAUnit (and lib ?) modules in the new Excel file project
    vtkImportModulesInAnotherProject projectForModules:=wb.VBProject, projectName:=project.projectName, confName:=project.projectDEVName
 
   ' Save configured and updated project for test
    wb.save
        
    ' Initialize git
    On Error GoTo vtkCreateProject_ErrorGit
    vtkInitializeGit rootPath

    On Error GoTo 0
    vtkCreateProject = VTK_OK
    Exit Function

vtkCreateProject_ErrorTreeFolder:
    vtkCreateProject = internalError
    If displayError Then MsgBox "Error " & err.number & " (" & err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
    Exit Function
vtkCreateProject_Error:
    vtkCreateProject = err.number
    If displayError Then MsgBox "Error " & err.number & " (" & err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
vtkCreateProject_ErrorGit:
    vtkCreateProject = err.number
    If displayError Then MsgBox "Error " & err.number & " (" & err.Description & ") in procedure vtkCreateProject of Module MainFunctions"

End Function
