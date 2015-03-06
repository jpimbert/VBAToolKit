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

    ' Create the vtkProject object attached to the new project
    Dim project As vtkProject
    Set project = vtkProjectForName(projectName:=name)
    Dim rootPath As String
    rootPath = path & "\" & project.projectName
    
    ' Create tree folder
    Dim internalError As Long
    internalError = vtkCreateTreeFolder(rootPath)
    If internalError <> VTK_OK Then GoTo vtkCreateProject_ErrorTreeFolder
     
    ' Create DEV workbook with xlsm extension
    Workbooks.Add.SaveAs (rootPath & "\" & project.projectDEVStandardRelativePath), FileFormat:=xlOpenXMLWorkbookMacroEnabled
    ' Rename the VBProject of the DEV workbook
    Workbooks(project.workbookDEVName).VBProject.name = project.projectDEVName
    ' Activate references and reference to the current workbook (the VBAToolKit add-in)
    VtkActivateReferences Wb:=Workbooks(project.workbookDEVName), projectName:=project.projectName, confName:=project.projectDEVName
    ' Initialize configuration Sheet with VBAUnit modules
    vtkInitializeVbaUnitNamesAndPathes project:=project.projectName
    ' Save DEV Workbook
    Workbooks(project.workbookDEVName).Save
    
    
    ' Create Delivery workbook with xlsm extension
    Workbooks.Add.SaveAs (rootPath & "\" & project.projectStandardRelativePath), FileFormat:=(52) '52 is xlsm format
    ' Rename the VBProject of the Delivery workbook
    Workbooks(project.workbookName).VBProject.name = project.projectName
    ' Activate references
    VtkActivateReferences Wb:=Workbooks(project.workbookName), projectName:=project.projectName, confName:=project.projectName
    ' A module must be added in the Excel File for the project parameters to be saved
    Workbooks(project.workbookName).VBProject.VBComponents.Add ComponentType:=vbext_ct_StdModule
    ' Save and Close Delivery Workbook
    Workbooks(project.workbookName).Close saveChanges:=True
    
    Dim Wb As Workbook
    Set Wb = Workbooks(project.workbookDEVName)
    Wb.Activate
    ' Get VBAUnit modules from VBAToolkit (This workbook = current running code)
    vtkExportModulesFromAnotherProject projectWithModules:=ThisWorkbook.VBProject, projectName:=project.projectName, confName:=project.projectDEVName
    ' Import VBAUnit (and lib ?) modules in the new Excel file project
    vtkImportModulesInAnotherProject projectForModules:=Wb.VBProject, projectName:=project.projectName, confName:=project.projectDEVName
    
    ' Insert the BeforeSave handler in the newly created project
    vtkAddBeforeSaveHandlerInDEVWorkbook Wb:=wb, projectName:=project.projectName, confName:=project.projectDEVName
    ' Declare the BeforeSave handler in the new project configuration
    Dim module As VBComponent, nm As Integer, nc As Integer, moduleName As String, cm As vtkConfigurationManager
    Set cm = vtkConfigurationManagerForProject(project.projectName)
    nc = cm.getConfigurationNumber(project.projectDEVName)
    moduleName = "thisWorkbook"
    Set module = Wb.VBProject.VBComponents(moduleName)
    nm = cm.addModule(moduleName)
    cm.setModulePathWithNumber path:="Source\ConfTest\" & module.name & ".cls", numModule:=nm, numConfiguration:=nc
    ' Save configured and updated project for test
    Wb.Save
        
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
' Procedure  : vtkRecreateConfigurations
' Author     : Jean-Pierre Imbert
' Date       : 15/07/2014
' Purpose    : recreate one or several configurations
' Parameters :
'             - confManager, configuration manager for the configurations to recreate
'             - confNames, Colelction of the name of the configurations to recreate
'
'---------------------------------------------------------------------------------------
'
Public Sub vtkRecreateConfigurations(confManager As vtkConfigurationManager, confNames As Collection)
    Dim c As Variant, confName As String
    For Each c In confNames
        confName = c
        vtkRecreateConfiguration confManager.projectName, confName, confManager
    Next c
End Sub
