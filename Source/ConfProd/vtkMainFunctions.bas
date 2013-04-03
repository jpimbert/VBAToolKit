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

'    Debug.Print CurDir
    
   On Error GoTo 0
   vtkCreateProject = 0
   Exit Function

vtkCreateProject_Error:
    vtkCreateProject = Err.Number
    If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
End Function
