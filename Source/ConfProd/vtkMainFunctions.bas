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
Public Function vtkCreateProject(Path As String, name As String, Optional displayError As Boolean = True) As Long
    
   
        'Application.DisplayAlerts = False 'to not display message that ask to save project
       On Error GoTo vtkCreateProject_Error

            ' Create main folder
            MkDir Path & "\" & name
            ' Create Delivery folder
            MkDir Path & "\" & name & "\" & "Delivery"
            ' Create Project folder
            MkDir Path & "\" & name & "\" & "Project"
            ' Create Tests folder
            MkDir Path & "\" & name & "\" & "Tests"
            ' Create GitLog Folder
            MkDir Path & "\" & name & "\" & "GitLog"
            ' Create Source folder
            MkDir Path & "\" & name & "\" & "Source"
            ' Create ConfProd folder
            MkDir Path & "\" & name & "\" & "Source" & "\" & "ConfProd"
            ' Create ConfTest folder
            MkDir Path & "\" & name & "\" & "Source" & "\" & "ConfTest"
            ' Create VbaUnit folder
            MkDir Path & "\" & name & "\" & "Source" & "\" & "VbaUnit"
             
           
             
            'Save created project with xlsm extention
            Workbooks.Add.SaveAs (Path & "\" & name & "\" & "Project" & "\" & name & "_Dev"), FileFormat:=(52) '52 is xlsm format
            'Rename Project
            Workbooks(name & "_Dev" & ".xlsm").VBProject.name = name & "_DEV"
            'call function who activate references
            VtkActivateReferences (name & "_Dev" & ".xlsm")
            'Create delivery workbook
            Workbooks.Add.SaveAs (Path & "\" & name & "\" & "Delivery" & "\" & name), FileFormat:=(52) '52 is xlsm format
            'Rename Project
            Workbooks(name & ".xlsm").VBProject.name = name
            
            'sheet name and paths will be implemented in vtkimportexportutilities
            vtkImportExportUtilities.DelivwbFullPAth = Workbooks(name & ".xlsm").FullNameURLEncoded
            vtkImportExportUtilities.DelivwbName = Workbooks(name & ".xlsm").name
            'activate dev workbook
            Workbooks(name & "_Dev" & ".xlsm").Activate
            'close delivery workbook "desactivate for tests"
            'Workbooks(name & ".xlsm").Close
            vtkImportExportUtilities.DevwbFullPAth = Workbooks(name & "_Dev" & ".xlsm").FullNameURLEncoded
            vtkImportExportUtilities.DevwbName = Workbooks(name & "_Dev" & ".xlsm").name
            vtkImportExportUtilities.DevWbFullName = Workbooks(name & "_Dev" & ".xlsm").VBProject.name
            
          retval2 = VtkInitilizeSheet()
    On Error GoTo 0
    vtkCreateProject = 0
    Exit Function
vtkCreateProject_Error:
    vtkCreateProject = Err.Number
If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
End Function

