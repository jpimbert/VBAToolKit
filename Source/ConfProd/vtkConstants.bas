Attribute VB_Name = "vtkConstants"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkConstants
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   :
'
' Copyright 2013 Skwal-Soft (http://skwalsoft.com)
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
'---------------------------------------------------------------------------------------

Public Const vtkTestProjectName = "TestProject"
Public Const vtkTestProjectNameWithExtention = "TestProject.xlsm"

Public Const VTK_OK = 0
Public Const VTK_UNEXPECTED_ERROR = 2001
Public Const VTK_WRONG_FOLDER_PATH = 2076
Public Const VTK_FORBIDDEN_PARAMETER = 2077

Public Const VTK_GIT_NOT_INSTALLED = 3000
Public Const VTK_GIT_ALREADY_INITIALIZED_IN_FOLDER = 3001
Public Const VTK_GIT_PROBLEM_DURING_INITIALIZATION = 3003

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
