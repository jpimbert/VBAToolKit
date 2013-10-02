Attribute VB_Name = "vtkFileSystemUtilities"
'---------------------------------------------------------------------------------------
' Module    : vtkFileSystemUtilities
' Author    : Lucas Vitorino
' Purpose   : Provide some utilities for interacting with files and folders.
'               - creation
'               - existence
'               - reading
'               - deletion...
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
' Procedure : vtkTextFileReader
' Author    : Abdelfattah Lahbib
' Date      : 30/04/2013
' Purpose   : Returns the content of a text file
' Notes     : Notably used to read Git log files.
'---------------------------------------------------------------------------------------
'
Public Function vtkTextFileReader(fullFilePath As String) As String

    Dim Textfile As Variant
    Dim strresult As String
    Dim fso As New FileSystemObject

On Error GoTo vtkTextFileReader_Error

    Set Textfile = fso.OpenTextFile(fullFilePath, ForReading)
    'while not end of file
    Do Until Textfile.AtEndOfStream
    'read line per line
        strresult = strresult & Chr(10) & Textfile.ReadLine
    Loop
    'return file text
    vtkTextFileReader = strresult

   On Error GoTo 0
   Exit Function

vtkTextFileReader_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure VtkTextFileReader of Module vtkGitFunctions"
    vtkTextFileReader = Err.Number
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkCleanFolder
' Author    : Lucas Vitorino
' Purpose   : Recursively delete all the content of a folder, leaving it empty.
' Notes     : Returns
'               - VTK_RETVAL_OK if successful
'               - 76 if wrong path or parameter is not a folder
'               - VTK_RETVAL_UNEXPECTED_ERR if other error
'---------------------------------------------------------------------------------------
'
Public Function vtkCleanFolder(folderPath As String) As Integer
    
    On Error GoTo vtkCleanFolder_Error
    
    Dim fso As New Scripting.FileSystemObject
    Dim sourceFolder As Scripting.Folder
    Dim subFolder As Scripting.Folder
    Dim File As Scripting.File
    ' Will raise an error if folderPath does not correspond to a valid folder
    Set sourceFolder = fso.GetFolder(folderPath)

    ' Erase the files in the folder, even the hidden ones
    For Each File In sourceFolder.Files
        fso.DeleteFile File
    Next File
    ' Call the function on all the SubFolders
    For Each subFolder In sourceFolder.SubFolders
        vtkCleanFolder (subFolder.path)
        fso.DeleteFolder subFolder.path, True
    Next subFolder
    
    On Error GoTo 0
    vtkCleanFolder = VTK_OK
    Exit Function
    
vtkCleanFolder_Error:
    If Err.Number = 76 Then
        vtkCleanFolder = Err.Number
    Else
        vtkCleanFolder = VTK_UNEXPECTED_ERROR
    End If
    Exit Function
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkDeleteFolder
' Author    : Lucas Vitorino
' Purpose   : Delete a folder and its content.
' Notes     : Returns
'               - VTK_RETVAL_OK if successful
'               - 76 if wrong path or parameter is not a folder
'               - VTK_RETVAL_UNEXPECTED_ERR if other error
'---------------------------------------------------------------------------------------
'
Public Function vtkDeleteFolder(folderPath As String)
    
    Dim fso As New Scripting.FileSystemObject
    Dim sourceFolder As Scripting.Folder

    On Error GoTo vtkDeleteFolder_Error

    'Will raise an error if the folder doesn't exist
    Set sourceFolder = fso.GetFolder(folderPath)
    
    vtkCleanFolder (folderPath)
    fso.DeleteFolder (sourceFolder.path)

    On Error GoTo 0
    vtkDeleteFolder = VTK_OK
    Exit Function

vtkDeleteFolder_Error:
    If Err.Number = 76 Then
        vtkDeleteFolder = Err.Number
    Else
        vtkDeleteFolder = VTK_UNEXPECTED_ERROR
    End If
    Exit Function

End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkDoesFolderExist
' Author    : Lucas Vitorino
' Purpose   : Checks if a folder exists.
' Returns   : Boolean. True if the folder exists, hidden or not, False in other cases.
'---------------------------------------------------------------------------------------
'
Public Function vtkDoesFolderExist(folderPath As String) As Integer

    On Error GoTo vtkDoesFolderExist_Error
    
    'Dir(etc,vbDirectory) returns True even if the specified thing is not a directory
    Dim fso As New FileSystemObject
    'Will raise an error 76 if wrong path or not a folder
    fso.GetFolder (folderPath)
    
    On Error GoTo 0
    vtkDoesFolderExist = True
    Exit Function

vtkDoesFolderExist_Error:
    vtkDoesFolderExist = False
    Exit Function

End Function

