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

'---------------------------------------------------------------------------------------
' Procedure : createxlsfile
' Author    : user
' Date      : 17/04/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function createxlsfile(path As String, name As String, Optional displayError As Boolean = True) As Long
    
    Dim Wb As Workbook
    Set Wb = Workbooks.Add
        
  
        On Error GoTo createxlsfile_Error
     'create an empty xls project
        Wb.SaveAs filename:=path & "\" & name & "\" & "Project" & "\" & name & ".xls"
       ' close created workbook
      ' Workbooks(name & ".xls").Close savechanges:=False ' we can't export modules when workbook is closed
        
        On Error GoTo 0
        createxlsfile = 0
        Exit Function
createxlsfile_Error:
        createxlsfile = Err.Number
        If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createxlsfile of Module MainFunctions"
      
End Function

'---------------------------------------------------------------------------------------
' Procedure : exportvbaunitwithproject
' Author    : user
' Date      : 17/04/2013
' Purpose   : -
'---------------------------------------------------------------------------------------
'
Function exportvbaunitwithproject(path As String, name As String) As Long

    Dim filename As String
    Dim i As Integer
    i = 0
    
    ChDir (ThisWorkbook.path)                                          'the current workbookpath
    ChDir ".."                                                         'allow acces to parent folder path
    vbaunitsourcepath = CurDir(ThisWorkbook.path) & "\Source\VbaUnit\" ' the vbaunitfolder path
   
  'init file
    filename = Dir(vbaunitsourcepath, vbNormal) 'DIR function returns the first filename vbNormal= default
  
  While filename <> ""
    'On Error Resume Next
    Workbooks(name & ".xls").VBProject.VBComponents.Import (vbaunitsourcepath & filename) 'add classes to new workbook
    FileCopy vbaunitsourcepath & filename, path & "\" & name & "\Source\VbaUnit\" & filename 'copy vbaunit file to destination directory
     filename = Dir
     i = i + 1
 Wend

exportvbaunitwithproject = i
End Function


