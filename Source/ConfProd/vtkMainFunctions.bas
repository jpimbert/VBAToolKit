Attribute VB_Name = "vtkMainFunctions"
'---------------------------------------------------------------------------------------
' Procedure : vtkCreateProject
' Author    : JPI-Conseil
' Date      : 03/04/2013
' Purpose   : Create a tree folder for a new project
'               - Source containing ConfProd, ConfTest and VBAUnit
'               - Project containing the main Excel file for the project
'               - an empty Tests folder
'               - A Git repository is initialized for the project
'               - Create Xlsm Dev and Delivery project
'               - Rename 2 Project
'               - Activate missing References
' Return    : Long error number
'
'   L'extension "Microsoft Visual Basic For Application Extensibility" doit être activée
'
'   some problems : -you can't create 2 project with the same name
'
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateProject(path As String, name As String, Optional displayError As Boolean = True) As Long
    
   
        'Application.DisplayAlerts = False 'to not display message that ask to save project
       On Error GoTo vtkCreateProject_Error

            ' Create the vtkProject attached to the new project
            Dim project As vtkProject
            Set project = vtkProjectForName(projectName:=name)
            
            ' Create main folder
            MkDir path & "\" & project.projectName
            ' Create Delivery folder
            MkDir path & "\" & project.projectName & "\" & "Delivery"
            ' Create Project folder
            MkDir path & "\" & project.projectName & "\" & "Project"
            ' Create Tests folder
            MkDir path & "\" & project.projectName & "\" & "Tests"
            ' Create GitLog Folder
            MkDir path & "\" & project.projectName & "\" & "GitLog"
            ' Create Source folder
            MkDir path & "\" & project.projectName & "\" & "Source"
            ' Create ConfProd folder
            MkDir path & "\" & project.projectName & "\" & "Source" & "\" & "ConfProd"
            ' Create ConfTest folder
            MkDir path & "\" & project.projectName & "\" & "Source" & "\" & "ConfTest"
            ' Create VbaUnit folder
            MkDir path & "\" & project.projectName & "\" & "Source" & "\" & "VbaUnit"
             
            'Save created project with xlsm extention
             Workbooks.Add.SaveAs (path & "\" & project.projectName & "\" & project.projectDEVStandardRelativePath), FileFormat:=(52) '52 is xlsm format
            'Rename Project
            Workbooks(project.projectDEVName & ".xlsm").VBProject.name = project.projectDEVName
            'call function who activate references
            VtkActivateReferences (project.projectDEVName)
            'initialize confsheet with dev workbook name and path
            Workbooks(project.projectDEVName & ".xlsm").Sheets(vtkConfSheet).Range(vtkModuleDevRange & vtkFirstLine - 2) = Workbooks(name & ".xlsm").FullNameURLEncoded
            Workbooks(project.projectDEVName & ".xlsm").Sheets(vtkConfSheet).Range(vtkModuleDevRange & vtkFirstLine - 3) = Workbooks(name & ".xlsm").name
            
            'Create delivery workbook
            Workbooks.Add.SaveAs (path & "\" & name & "\" & project.projectStandardRelativePath), FileFormat:=(52) '52 is xlsm format
            'Rename Project
            Workbooks(project.projectName & ".xlsm").VBProject.name = project.projectName
            'call function who activate references
            VtkActivateReferences (project.projectName)
            'initialize confsheet with delivery workbook name and path
            Workbooks(project.projectName & ".xlsm").Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & vtkFirstLine - 2) = Workbooks(name & "_Delivery" & ".xlsm").FullNameURLEncoded
            Workbooks(project.projectName & ".xlsm").Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & vtkFirstLine - 3) = Workbooks(name & "_Delivery" & ".xlsm").name
            'activate dev workbook
            Workbooks(project.projectDEVName).Activate
            '
            RetVtkExportAll = vtkExportAll(ThisWorkbook.name)
            RetValImportTestConf = vtkImportTestConfig()
    On Error GoTo 0
    vtkCreateProject = 0
    Exit Function
vtkCreateProject_Error:
    vtkCreateProject = Err.Number
If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
End Function
