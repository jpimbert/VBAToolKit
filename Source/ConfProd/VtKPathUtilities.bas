Attribute VB_Name = "VtKPathUtilities"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkPathUtilities
' Author    : Jean-Pierre Imbert
' Date      : 03/07/2013
' Purpose   : This module contains utility fonctions for obtaining various folder
'             pathes of the project.
'
'             This module is primarily used within VBAToolKit unit tests
'             It could be duplicated in projects managed with VBAToolKit for Unit Tests of these projects
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : vtkPathOfCurrentProject
' Author    : Jean-Pierre Imbert
' Date      : 18/04/2013
' Purpose   : Return the path of the current project
'               - Application.ThisWorbook is the workbook containing the running code
'---------------------------------------------------------------------------------------
'
Public Function vtkPathOfCurrentProject() As String
    Dim fso As New FileSystemObject
    vtkPathOfCurrentProject = fso.GetParentFolderName(ThisWorkbook.path)
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkPathToTestFolder
' Author    : Jean-Pierre Imbert
' Date      : 18/04/2013
' Purpose   : Return the path of the Test Folder of the current project
'               - create the folder if it doesn't exist (in case of fresh Git check out)
'---------------------------------------------------------------------------------------

Public Function vtkPathToTestFolder() As String '\VBAToolKit\Tests
    Dim path As String
    path = vtkPathOfCurrentProject & "\Tests"
    If Dir(path, vbDirectory) = vbNullString Then MkDir (path)
    vtkPathToTestFolder = path
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkPathToSourceFolder
' Author    : Jean-Pierre Imbert
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



