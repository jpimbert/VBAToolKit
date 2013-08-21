Attribute VB_Name = "vtkMainFunctions"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkMainFunctions
' Author    : Jean-Pierre Imbert
' Date      : 04/07/2013
' Purpose   : This module contains the functions called for the main capacities of VBAToolKit
'               - new project creation
'               - (other capacities will be delelopped later)
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
    internalError = vtkCreateTreeFolderWithGitIgnore(rootPath)
    If internalError <> VTK_OK Then GoTo vtkCreateProject_ErrorTreeFolder
     
    'Save created project with xlsm extention
    Workbooks.Add.SaveAs (rootPath & "\" & project.projectDEVStandardRelativePath), FileFormat:=xlOpenXMLWorkbookMacroEnabled
    'Rename Project
    Workbooks(project.workbookDEVName).VBProject.name = project.projectDEVName
    'call function who activate references
    VtkActivateReferences (project.workbookDEVName)
    'initialize configuration Sheet with VBAUnit modules
    vtkInitializeVbaUnitNamesAndPathes project:=project.projectName
    ' Save Development Project Workbook
    Workbooks(project.workbookDEVName).save
    
    'Create delivery workbook
    Workbooks.Add.SaveAs (rootPath & "\" & project.projectStandardRelativePath), FileFormat:=(52) '52 is xlsm format
    'Rename Project
    Workbooks(project.workbookName).VBProject.name = project.projectName
    'call function who activate references
    VtkActivateReferences (project.workbookName)
    ' A module must be added in the Excel File for the project parameters to be saved
    Workbooks(project.workbookName).VBProject.VBComponents.Add ComponentType:=vbext_ct_StdModule
    ' Save and Close Delivery Project WorkBook
    Workbooks(project.workbookName).Close SaveChanges:=True
    
    Workbooks(project.workbookDEVName).Activate
    '
    '            RetVtkExportAll = vtkExportAll(ThisWorkbook.name)
    '            RetValImportTestConf = vtkImportTestConfig()
    On Error GoTo 0
    vtkCreateProject = 0
    Exit Function
vtkCreateProject_ErrorTreeFolder:
    vtkCreateProject = internalError
    If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
    Exit Function
vtkCreateProject_Error:
    vtkCreateProject = Err.Number
    If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
End Function
