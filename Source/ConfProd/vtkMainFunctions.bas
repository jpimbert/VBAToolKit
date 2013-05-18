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

            ' Create main folder
            MkDir path & "\" & name
            ' Create Delivery folder
            MkDir path & "\" & name & "\" & "Delivery"
            ' Create Project folder
            MkDir path & "\" & name & "\" & "Project"
            ' Create Tests folder
            MkDir path & "\" & name & "\" & "Tests"
            ' Create GitLog Folder
            MkDir path & "\" & name & "\" & "GitLog"
            ' Create Source folder
            MkDir path & "\" & name & "\" & "Source"
            ' Create ConfProd folder
            MkDir path & "\" & name & "\" & "Source" & "\" & "ConfProd"
            ' Create ConfTest folder
            MkDir path & "\" & name & "\" & "Source" & "\" & "ConfTest"
            ' Create VbaUnit folder
            MkDir path & "\" & name & "\" & "Source" & "\" & "VbaUnit"
             
            'Save created project with xlsm extention
             Workbooks.Add.SaveAs (path & "\" & name & "\" & "Project" & "\" & name), FileFormat:=(52) '52 is xlsm format
            'Rename Project
            Workbooks(name & ".xlsm").VBProject.name = name & "_DEV"
            'call function who activate references
            VtkActivateReferences (name & ".xlsm")
            'initialize confsheet with dev workbook name and path
            Workbooks(name & ".xlsm").Sheets(vtkConfSheet).Range(vtkModuleDevRange & vtkFirstLine - 2) = Workbooks(name & ".xlsm").FullNameURLEncoded
            Workbooks(name & ".xlsm").Sheets(vtkConfSheet).Range(vtkModuleDevRange & vtkFirstLine - 3) = Workbooks(name & ".xlsm").name
            
            'Create delivery workbook
            Workbooks.Add.SaveAs (path & "\" & name & "\" & "Delivery" & "\" & name & "_Delivery"), FileFormat:=(52) '52 is xlsm format
            'Rename Project
            Workbooks(name & "_Delivery" & ".xlsm").VBProject.name = name & "_Delivery"
            'call function who activate references
            VtkActivateReferences (name & "_Delivery" & ".xlsm")
            'initialize confsheet with delivery workbook name and path
            Workbooks(name & ".xlsm").Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & vtkFirstLine - 2) = Workbooks(name & "_Delivery" & ".xlsm").FullNameURLEncoded
            Workbooks(name & ".xlsm").Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & vtkFirstLine - 3) = Workbooks(name & "_Delivery" & ".xlsm").name
            'activate dev workbook
            Workbooks(name & ".xlsm").Activate
            '
            a = vtkExportAll(ThisWorkbook.name)
    On Error GoTo 0
    vtkCreateProject = 0
    Exit Function
vtkCreateProject_Error:
    vtkCreateProject = Err.Number
If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
End Function
