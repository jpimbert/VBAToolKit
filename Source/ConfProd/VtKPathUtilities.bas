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

'---------------------------------------------------------------------------------------
' Procedure : vtkGetFileExtension
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Return the extension of the file whose path is given as parameter
'             - return "" is the filepath has no extension
'---------------------------------------------------------------------------------------
'
Public Function vtkGetFileExtension(filePath As String) As String
    Dim dotPosition As Integer
    dotPosition = InStrRev(filePath, ".")
    If dotPosition = 0 Then
        vtkGetFileExtension = ""
       Else
        vtkGetFileExtension = Mid(filePath, dotPosition + 1)
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : vtkCreateTreeFolder
' Author    : Jean-Pierre Imbert
' Date      : 06/08/2013
' Purpose   : Create a project folder breakdown into the folder given as parameter
'             This procedure is isolated to be easier to test
' Return    : Long error number
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateTreeFolder(rootPath As String)
   On Error GoTo vtkCreateTreeFolder_Error
    
    MkDir rootPath
    MkDir rootPath & "\" & "Delivery"
    MkDir rootPath & "\" & "Project"
    MkDir rootPath & "\" & "Tests"
    MkDir rootPath & "\" & "GitLog"
    MkDir rootPath & "\" & "Source"
    MkDir rootPath & "\" & "Source" & "\" & "ConfProd"
    MkDir rootPath & "\" & "Source" & "\" & "ConfTest"
    MkDir rootPath & "\" & "Source" & "\" & "VbaUnit"

   On Error GoTo 0
   vtkCreateTreeFolder = VTK_OK
   Exit Function
   
vtkCreateTreeFolder_Error:
    vtkCreateTreeFolder = err.Number
    err.Raise err.Number, "Module vtkFileSystemUtilities : Function vtkCreateTreeFolder", err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkDeleteTreeFolder
' Author    : Jean-Pierre Imbert
' Date      : 06/08/2013
' Purpose   : Delete a project folder breakdown given as parameter
'             This procedure is for test purpose
'---------------------------------------------------------------------------------------
'
Public Sub vtkDeleteTreeFolder(rootPath As String)
    Dir (rootPath)                  ' Make sure to be out of the folder to clean it without Err
    On Error Resume Next
    Kill rootPath & "\Source\ConfProd\*"
    RmDir rootPath & "\Source\ConfProd"
    Kill rootPath & "\Source\ConfTest\*"
    RmDir rootPath & "\Source\ConfTest"
    Kill rootPath & "\Source\VbaUnit\*"
    RmDir rootPath & "\Source\VbaUnit"
    Kill rootPath & "\GitLog\*"
    RmDir rootPath & "\GitLog"
    Kill rootPath & "\Tests\*"
    RmDir rootPath & "\Tests"
    Kill rootPath & "\Source\*"
    RmDir rootPath & "\Source"
    Kill rootPath & "\Delivery\*"
    RmDir rootPath & "\Delivery"
    Kill rootPath & "\Project\*"
    RmDir rootPath & "\Project"
    Kill rootPath & "\*"
    RmDir rootPath
End Sub


