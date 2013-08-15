Attribute VB_Name = "vtkConstants"
Public Const vtkTestProjectName = "TestProject"
Public Const vtkTestProjectNameWithExtention = "TestProject.xlsm"
Public Const VTK_OK = "0"
Public Const VTK_UNEXPECTED_ERROR = "2000"
Public Const VTK_WRONG_FOLDER_PATH = "2076"
Public Const VTK_FORBIDDEN_PARAMETER = "2077"

Public Const VTK_GIT_NOT_INSTALLED = "3000"
Public Const VTK_GIT_ALREADY_INITIALIZED_IN_FOLDER = "3001"

'---------------------------------------------------------------------------------------
' Procedure : vtkTestPath
' Author    : Abdelfattah Lahbib
' Date      : 07/05/2013
' Purpose   : Return the path of the Test Folder of the current project  '..\VBAToolKit\Tests
'---------------------------------------------------------------------------------------
'
Public Function vtkTestPath() As String
    vtkTestPath = vtkPathToTestFolder
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInstallPath
' Author    : Abdelfattah Lahbib
' Date      : 07/05/2013
' Purpose   : Return the path of the current project  '..\VBAToolKit
'---------------------------------------------------------------------------------------
'
Public Function vtkInstallPath() As String
    vtkInstallPath = vtkPathOfCurrentProject
End Function
