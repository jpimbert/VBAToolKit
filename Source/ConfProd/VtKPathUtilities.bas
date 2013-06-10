Attribute VB_Name = "VtKPathUtilities"
'Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : vtkPathOfCurrentProject
' Author    : Demonn
' Date      : 18/04/2013
' Purpose   : Return the path of the current project
'---------------------------------------------------------------------------------------
'
Public Function vtkPathOfCurrentProject() As String
    Dim fso As New FileSystemObject
    vtkPathOfCurrentProject = fso.GetParentFolderName(ThisWorkbook.path)
End Function
Public Function testtest(projectName) As String
  
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkPathToTestFolder
' Author    : Demonn
' Date      : 18/04/2013
' Purpose   : Return the path of the Test Folder of the current project
'---------------------------------------------------------------------------------------

Public Function vtkPathToTestFolder() As String '\VBAToolKit\Tests
    
    vtkPathToTestFolder = vtkPathOfCurrentProject & "\Tests"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkPathToSourceFolder
' Author    : Demonn
' Date      : 18/04/2013
' Purpose   : Return the path of the Source Folder of the current project
'---------------------------------------------------------------------------------------
'
Public Function vtkPathToSourceFolder() As String 'VBAToolKit\Source

   vtkPathToSourceFolder = vtkPathOfCurrentProject & "\Source"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkPathToTemplateFolder
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Return the path of the Template Folder of the current project
'---------------------------------------------------------------------------------------
'
Public Function vtkPathToTemplateFolder() As String 'VBAToolKit\Source

   vtkPathToTemplateFolder = vtkPathOfCurrentProject & "\Templates"
End Function



