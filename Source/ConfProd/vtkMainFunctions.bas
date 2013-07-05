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

    Dim rootPath As String
    rootPath = path & "\" & project.projectName
    
    ' Create the vtkProject object attached to the new project
    Dim project As vtkProject
    Set project = vtkProjectForName(projectName:=name)
    
    ' Create main folder
    MkDir rootPath
    ' Create Delivery folder
    MkDir rootPath & "\" & "Delivery"
    ' Create Project folder
    MkDir rootPath & "\" & "Project"
    ' Create Tests folder
    MkDir rootPath & "\" & "Tests"
    ' Create GitLog Folder
    MkDir rootPath & "\" & "GitLog"
    ' Create Source folder
    MkDir rootPath & "\" & "Source"
    ' Create ConfProd folder
    MkDir rootPath & "\" & "Source" & "\" & "ConfProd"
    ' Create ConfTest folder
    MkDir rootPath & "\" & "Source" & "\" & "ConfTest"
    ' Create VbaUnit folder
    MkDir rootPath & "\" & "Source" & "\" & "VbaUnit"
     
    'Save created project with xlsm extention
    Workbooks.Add.SaveAs (rootPath & "\" & project.projectDEVStandardRelativePath), FileFormat:=xlOpenXMLWorkbookMacroEnabled
    'Rename Project
    Workbooks(project.workbookDEVName).VBProject.name = project.projectDEVName
    'call function who activate references
    VtkActivateReferences (project.workbookDEVName)
    'initialize configuration Sheet with VBAUnit modules
    vtkInitializeVbaUnitNamesAndPathes project:=project.projectName
    ' Save Development Project Workbook
    Workbooks(project.workbookDEVName).Save
    
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
vtkCreateProject_Error:
    vtkCreateProject = Err.Number
    If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeVbaUnitNamesAndPathes
' Author    : Abdelfattah Lahbib
' Date      : 09/05/2013
' Purpose   : - Initialize DEV project ConfSheet with vbaunit module names and pathes
'             - Return True if module names and paths are initialized without error
'---------------------------------------------------------------------------------------
'
Public Function vtkInitializeVbaUnitNamesAndPathes(project As String) As Boolean
    Dim tableofvbaunitname(17) As String
        tableofvbaunitname(0) = "VbaUnitMain"
        tableofvbaunitname(1) = "Assert"
        tableofvbaunitname(2) = "AutoGen"
        tableofvbaunitname(3) = "IAssert"
        tableofvbaunitname(4) = "IResultUser"
        tableofvbaunitname(5) = "IRunManager"
        tableofvbaunitname(6) = "ITest"
        tableofvbaunitname(7) = "ITestCase"
        tableofvbaunitname(8) = "ITestManager"
        tableofvbaunitname(9) = "RunManager"
        tableofvbaunitname(10) = "TestCaseManager"
        tableofvbaunitname(11) = "TestClassLister"
        tableofvbaunitname(12) = "TesterTemplate"
        tableofvbaunitname(13) = "TestFailure"
        tableofvbaunitname(14) = "TestResult"
        tableofvbaunitname(15) = "TestRunner"
        tableofvbaunitname(16) = "TestSuite"
        tableofvbaunitname(17) = "TestSuiteManager"
    Dim i As Integer, cm As vtkConfigurationManager, ret As Boolean, nm As Integer, nc As Integer, ext As String
    Set cm = vtkConfigurationManagerForProject(project)
    nc = cm.getConfigurationNumber(vtkProjectForName(project).projectDEVName)
    ret = (nc > 0)
    For i = LBound(tableofvbaunitname) To UBound(tableofvbaunitname)
        nm = cm.addModule(tableofvbaunitname(i))
        ret = ret And (nm > 0)
        If i <= 0 Then      ' It's a Standard Module (WARNING, magical number)
            ext = ".bas"
           Else
            ext = ".cls"    ' It's a Class Module
        End If
        cm.setModulePathWithNumber path:="Source\VbaUnit\" & tableofvbaunitname(i) & ext, numModule:=nm, numConfiguration:=nc
    Next i
    vtkInitializeVbaUnitNamesAndPathes = ret
End Function


