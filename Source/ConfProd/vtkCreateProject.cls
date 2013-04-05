VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkCreateProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkCreateProject
' Author    : user
' Date      : 04/04/2013
' Purpose   :
'---------------------------------------------------------------------------------------



Option Explicit




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
    
'added by lahbib
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


