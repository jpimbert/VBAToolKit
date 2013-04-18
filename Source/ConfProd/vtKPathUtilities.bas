Attribute VB_Name = "vtKPathUtilities"
Option Explicit

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

'---------------------------------------------------------------------------------------
' Procedure : vtkPathToTestFolder
' Author    : Demonn
' Date      : 18/04/2013
' Purpose   : Return the path of the Test Folder of the current project
'---------------------------------------------------------------------------------------
'
Public Function vtkPathToTestFolder() As String
    vtkPathToTestFolder = vtkPathOfCurrentProject & "\Tests"
End Function

Public Function vtkTestPath() As String
    vtkTestPath = vtkPathToTestFolder
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkPathToSourceFolder
' Author    : Demonn
' Date      : 18/04/2013
' Purpose   : Return the path of the Source Folder of the current project
'---------------------------------------------------------------------------------------
'
Public Function vtkPathToSourceFolder() As String
    vtkPathToSourceFolder = vtkPathOfCurrentProject & "\Source"
End Function


