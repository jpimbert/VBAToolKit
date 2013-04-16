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
' Return    : Boolean True if the project is created
'
'   L'extension "Microsoft Visual Basic For Application Extensibility" doit être activée
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateProject(path As String, name As String, Optional displayError As Boolean = True) As Long
    
   
    
        On Error GoTo vtkCreateProject_Error

            ' Create main folder
            MkDir path & "\" & name
            ' Create Project folder
            MkDir path & "\" & name & "\" & "Project"
            ' Create Tests folder
            MkDir path & "\" & name & "\" & "Tests"
            ' Create Source folder
            MkDir path & "\" & name & "\" & "Source"
            ' Create ConfProd folder
            MkDir path & "\" & name & "\" & "Source" & "\" & "ConfProd"
            ' Create ConfTest folder
            MkDir path & "\" & name & "\" & "Source" & "\" & "ConfTest"
            ' Create VbaUnit folder
            MkDir path & "\" & name & "\" & "Source" & "\" & "VbaUnit"
            
           
            'end added

'    Debug.Print CurDir
    
    On Error GoTo 0
    vtkCreateProject = 0
    Exit Function
vtkCreateProject_Error:
    vtkCreateProject = Err.Number
    If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
End Function

Public Function createxlsfile(path As String, name As String, Optional displayError As Boolean = True) As Long
    
        Dim Wb As Workbook
        Set Wb = Workbooks.Add
        Dim pathandfilename As String
        pathandfilename = path & "\" & name & "\" & "Project" & "\" & name & ".xls"
        On Error GoTo createxlsfile_Error
     'create an empty xls project
        Wb.SaveAs Filename:=pathandfilename
       ' close created workbook
        Workbooks(name & ".xls").Close savechanges:=False
        
        
    
        On Error GoTo 0
        createxlsfile = 0
        Exit Function
createxlsfile_Error:
        createxlsfile = Err.Number
        If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createxlsfile of Module MainFunctions"
      
End Function
