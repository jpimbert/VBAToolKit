Attribute VB_Name = "vtkConstants"
Public Const vtkTestProjectName = "TestProject"
Public Const vtkTestProjectNameWithExtention = "TestProject.xlsm"

'---------------------------------------------------------------------------------------
' Procedure : vtkTestPath
' Author    : Jean-Pierre Imbert (please confirm)
' Date      : 07/05/2013
' Purpose   : -Return the path of the Test Folder of the current project  '..\VBAToolKit\Tests
'---------------------------------------------------------------------------------------
'
Public Function vtkTestPath() As String
    vtkTestPath = vtkPathToTestFolder
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInstallPath
' Author    : Jean-Pierre Imbert (please confirm)
' Date      : 07/05/2013
' Purpose   :-Return the path of the current project  '..\VBAToolKit
'---------------------------------------------------------------------------------------
'
Public Function vtkInstallPath() As String
    vtkInstallPath = vtkPathOfCurrentProject
End Function
