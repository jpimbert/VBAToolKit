Attribute VB_Name = "vtkConstants"
Public Const vtkTestProjectName = "TestProject"
Public Const vtkTestProjectNameWithExtention = "TestProject.xlsm"

'---------------------------------------------------------------------------------------
' Procedure : vtkTestPath
' Author    : user
' Date      : 07/05/2013
' Purpose   : -Return the path of the Test Folder of the current project  '..\VBAToolKit\Tests
'---------------------------------------------------------------------------------------
'
Public Function vtkTestPath() As String
    vtkTestPath = vtkPathToTestFolder
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInstallPath
' Author    : user
' Date      : 07/05/2013
' Purpose   :-Return the path of the current project  '..\VBAToolKit
'            -Public Const vtkInstallPath = vtkPathOfCurrentProject : don't work , pas possibled'affecter une var a un const
'---------------------------------------------------------------------------------------
'
Public Function vtkInstallPath() As String
    vtkInstallPath = vtkPathOfCurrentProject
End Function
