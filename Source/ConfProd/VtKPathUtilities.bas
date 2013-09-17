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
'               - if a Workbook is given as parameter, return the root path of this project workbook
'---------------------------------------------------------------------------------------
'
Public Function vtkPathOfCurrentProject(Optional wb As Workbook) As String
    Dim fso As New FileSystemObject
    If wb Is Nothing Then
        vtkPathOfCurrentProject = fso.GetParentFolderName(ThisWorkbook.path)
       Else
        vtkPathOfCurrentProject = fso.GetParentFolderName(wb.path)
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkPathToTestFolder
' Author    : Jean-Pierre Imbert
' Date      : 18/04/2013
' Purpose   : Return the path of the Test Folder of the current project
'               - create the folder if it doesn't exist (in case of fresh Git check out)
'               - if a Workbook is given as parameter, return the test path of this project workbook
'---------------------------------------------------------------------------------------

Public Function vtkPathToTestFolder(Optional wb As Workbook) As String '\VBAToolKit\Tests
    Dim path As String
    path = vtkPathOfCurrentProject(wb) & "\Tests"
    If Dir(path, vbDirectory) = vbNullString Then MkDir (path)
    vtkPathToTestFolder = path
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkPathToSourceFolder
' Author    : Jean-Pierre Imbert
' Date      : 18/04/2013
' Purpose   : Return the path of the Source Folder of the current project
'               - if a Workbook is given as parameter, return the source path of this project workbook
'---------------------------------------------------------------------------------------
'
Public Function vtkPathToSourceFolder(Optional wb As Workbook) As String 'VBAToolKit\Source
   vtkPathToSourceFolder = vtkPathOfCurrentProject(wb) & "\Source"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkPathToTemplateFolder
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Return the path of the Template Folder of the current project
'               - if a Workbook is given as parameter, return the template path of this project workbook
'---------------------------------------------------------------------------------------
'
Public Function vtkPathToTemplateFolder(Optional wb As Workbook) As String 'VBAToolKit\Source
   vtkPathToTemplateFolder = vtkPathOfCurrentProject(wb) & "\Templates"
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
' Procedure : vtkStripPathOrNameOfExtension
' Author    : Lucas Vitorino
' Purpose   : - Returns a String containing the name of the file, that is to say without the first extension.
'             - Works with a path too.
' Examples  : "dummy.dummy" -> "dummy"
'             "dummy" -> "dummy"
'             "dummy\dummy" -> "dummy"
'             "dummy\dummy.dummy" -> "dummy"
'             "dummy.dummy.dummy" -> "dummy.dummy"
'---------------------------------------------------------------------------------------
'
Public Function vtkStripFilePathOrNameOfExtension(fileNameOrPath As String) As String

    On Error GoTo vtkStripFilePathOrNameOfExtension_Error

    Dim backslashPosition As Integer
    backslashPosition = InStrRev(fileNameOrPath, "\")
    
    ' Get the filename in the path
    Dim substring As String
    substring = Mid(fileNameOrPath, backslashPosition + 1)
    
    ' Remove the part after the last dot
    Dim dotPosition As Integer
    dotPosition = InStrRev(substring, ".")
    If dotPosition <> 0 Then
        substring = Left(substring, dotPosition - 1)
    End If
    
    vtkStripFilePathOrNameOfExtension = substring

    On Error GoTo 0
    Exit Function

vtkStripFilePathOrNameOfExtension_Error:
    Err.Raise VTK_UNEXPECTED_ERROR, "vtkStripFileNameOfExtension", Err.Description
    Resume Next

End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkStripVtkExtensionOfPathOrName
' Author    : Lucas Vitorino
' Purpose   : - Returns the vtkProjectName associated with a workbook path or name or a configuration name.
'
' Examples  : "Dummy_DEV", "DEV" -> "Dummy"
'             "Dummy_DEV.dummy", "DEV" -> "Dummy"
'             "Dummy.dummy" -> "Dummy"
'             "dummy\Dummy_DEV.dummy", "DEV" -> "Dummy"
'             "Dummy_DEV", "DEEV" -> "Dummy_DEV"
'             "Dummy_Dummy_DEV", "DEV", -> "Dummy_Dummy"
'             "Dummy_Dummy"_DEV", "Dummy" -> "Dummy_Dummy_DEV"
'---------------------------------------------------------------------------------------
'
Public Function vtkStripPathOrNameOfVtkExtension(projectNameOrPath As String, extension As String) As String

    On Error GoTo vtkStripPathOrNameOfVtkExtension_Error

    ' Get the name with or without the "_"
    Dim substring As String
    substring = vtkStripFilePathOrNameOfExtension(projectNameOrPath)
    
    ' Strip from the last "_" if the part after corresponds to the specified extension
    Dim underscorePosition As Integer
    underscorePosition = InStrRev(substring, "_")
    If Mid(substring, underscorePosition + 1) Like extension Then
        substring = Left(substring, underscorePosition - 1)
    End If
    
    vtkStripPathOrNameOfVtkExtension = substring

    On Error GoTo 0
    Exit Function

vtkStripPathOrNameOfVtkExtension_Error:
    Err.Raise VTK_UNEXPECTED_ERROR, "vtkStripPathOrNameOfVtkExtension", Err.Description
    Resume Next

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
    vtkCreateTreeFolder = Err.number
    Err.Raise Err.number, "Module vtkFileSystemUtilities : Function vtkCreateTreeFolder", Err.Description
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


