Attribute VB_Name = "vtkGitFunctions"
'---------------------------------------------------------------------------------------
' Module    : vtkGitFunctions
' Author    : Lucas Vitorino
' Purpose   : Interact with Git
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

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeGit
' Author    : Lucas Vitorino
' Purpose   : - Initialize Git in a directory using the "git init" command.
'               The directory path has to be absolute, and on the C: drive
'             - Optionally, writes the output of the "git init" command in a textfile
'               whose path has to be passed as parameter.
'               If the path is absolute, a drive other than C: will raise an error
'               If the path is relative, it will be considered as relatie to the folder path.
'             - Delete the previous $GIT_FOLDER/info/exclude file and creates a new one
'               that excludes :
'               - the content of the $VTK_PROJECT_FOLDER/Tests
'               - the content of the $VTK_PROJECT_FOLDER/GitLog
'               - the temporary files in $VTK_PROJECT_FOLDER/Project
'               - the Excel files in $VTK_PROJECT_FOLDER/Delivery
'             - Adds all the files in the directory to the git repository.
' Notes     : Raise errors
'               - VTK_WRONG_FOLDER_PATH
'               - VTK_FORBIDDEN_PARAMETER
'               - VTK_GIT_NOT_INSTALLED
'               - VTK_GIT_ALREADY_INITIALIZED_IN_FOLDER
'               - VTK_GIT_PROBLEM_DURING_INITIALIZATION
'               - VTK_UNEXPECTED_ERR
'---------------------------------------------------------------------------------------
'
Public Function vtkInitializeGit(folderPath As String, Optional logFile As String = "")
    Dim tmpLogFileName As String
    tmpLogFileName = "initialize.log"
    Dim logFileFullPath As String
    
    Dim convertedFolderPath As String
    Dim convertedLogFilePath As String
    
    Dim fso As New FileSystemObject
    Dim contentStream As TextStream
    
    On Error GoTo vtkInitializeGit_Err
        
    If InStr(UCase(Environ("PATH")), UCase("Git\cmd")) = False Then
        Err.Raise VTK_GIT_NOT_INSTALLED, "", "Git not installed."
    End If
    
    ' Potentially raise VTK_FORBIDDEN_PARAMETER
    convertedFolderPath = vtkGitConvertWinPath(folderPath)
    
    If vtkDoesFolderExist(folderPath) = False Then
        Err.Raise VTK_WRONG_FOLDER_PATH, "", "Folder path not found."
    End If
    
    If vtkDoesFolderExist(folderPath & "\.git") = True Then
        Err.Raise VTK_GIT_ALREADY_INITIALIZED_IN_FOLDER, "", "Git has already been initialized in the folder " & folderPath
    End If
    
    ' Get the path of the log file that will be used
    ' If a log file has been passed as a parameter
    If logFile <> "" Then
        ' If the path is relative, make it absolute
        Dim splittedLogFilePath() As String
        splittedLogFilePath = Split(logFile, ":")
        If splittedLogFilePath(LBound(splittedLogFilePath)) = logFile Then
            logFileFullPath = folderPath & "\" & logFile
        End If
        ' convert and potentially raise error
        convertedLogFilePath = vtkGitConvertWinPath(logFileFullPath)
    Else
        convertedLogFilePath = vtkGitConvertWinPath(folderPath & "\" & tmpLogFileName)
    End If
    
    ' Intializing git using a shell command and redirecting the output flow in the log file
    ShellAndWait "cmd.exe /c git init " & convertedFolderPath _
    & " > " & convertedLogFilePath, 0, vbHide, AbandonWait
    
    ' Check if the initialization went well
    Dim logFileContent As String
    logFileContent = vtkTextFileReader(folderPath & "\" & tmpLogFileName)
    If Left(logFileContent, 12) <> Chr(10) & "Initialized" Then
        Err.Raise VTK_GIT_PROBLEM_DURING_INITIALIZATION, , "There was a problem during Git initialization." _
        & vbCrLf & "Content of the log file : " & logFileContent
    End If
    
    ' Delete file if tmp
    If logFile = "" Then
        Kill folderPath & "\" & tmpLogFileName
    End If
    
    ' Delete the default git exclude file
    fso.DeleteFile folderPath & "\.git\info\exclude", True
    
    ' Create it again and fill it with the content we want.
    Set contentStream = fso.CreateTextFile(folderPath & "\.git\info\exclude")
    contentStream.WriteLine "# Ignore the content of the Tests and GitLog folders"
    contentStream.WriteLine "/Tests/*"
    contentStream.WriteLine "/GitLog/*"
    contentStream.WriteLine
    contentStream.WriteLine "# Ignore the temporary Excel files"
    contentStream.WriteLine "~*"
    contentStream.WriteLine
    contentStream.WriteLine "# Ignore the delivery Excel files"
    contentStream.WriteLine "/Delivery/*.xl*"
    contentStream.WriteLine
    contentStream.WriteLine "# Ignore the Project Excel files"
    contentStream.WriteLine "/Project/*.xl*"
    contentStream.Close
    
    ' Adds all the files in the folder tree to the git repository
    ShellAndWait "cmd.exe /c cd " & folderPath & " & git add " & ". ", 0, vbHide, AbandonWait
     
    On Error GoTo 0
    vtkInitializeGit = VTK_OK
    Exit Function
    
    
vtkInitializeGit_Err:
    If ((Err.Number = VTK_GIT_NOT_INSTALLED) _
        Or (Err.Number = VTK_GIT_ALREADY_INITIALIZED_IN_FOLDER) _
        Or (Err.Number = VTK_FORBIDDEN_PARAMETER) _
        Or (Err.Number = VTK_GIT_PROBLEM_DURING_INITIALIZATION) _
        Or (Err.Number = VTK_WRONG_FOLDER_PATH)) Then
        Err.Raise Err.Number, "Module vktGitFuntions : Function vtkGitInitialize", Err.Description
    Else
        Debug.Print "ERR IN INITIALIZE : " & Err.Number & Err.Description
        Err.Raise VTK_UNEXPECTED_ERROR, "Module vktGitFuntions : Function vtkGitInitialize", Err.Description
    End If
     
    Exit Function

End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkGitConvertWinPath
' Author    : Lucas Vitorino
' Purpose   : Converts an *absolute* Windows PATH to a one suitable for use with Git.
' Notes     : - WIP because I don't fully understand the behaviour of Git. I suppose
'               it's a Windows/Unix path format conflict.
'             - For now, I can't specify the drive letter in a path used with Git.
'               This function will strip it and git will assume it's "C" .
'             - Returns
'               - Converted string if OK
'               - VTK_FORBIDDEN_PARAMETER if winPath is not absolute, or absolute but not on the C: drive
'               - VTK_UNEXPECTED_ERR
'---------------------------------------------------------------------------------------
'
Public Function vtkGitConvertWinPath(winPath As String) As String
    
    Dim convertedPath As String
    convertedPath = winPath
    Dim convertedSplittedPath() As String
    
    On Error GoTo vtkGitConvertWinPath_Error
    
    convertedSplittedPath = Split(convertedPath, ":")
    
    ' Only allows absolute paths on the C: drive
    If convertedSplittedPath(LBound(convertedSplittedPath)) <> "C" Then
        Err.Raise VTK_FORBIDDEN_PARAMETER, "", "Parameter is invalid."
    End If

    convertedPath = convertedSplittedPath(LBound(convertedSplittedPath) + 1)
    
    'Changing the backslahes in slashes
    ' NB : Optional
    convertedPath = Replace(convertedPath, "\", "/")
        
    On Error GoTo 0
    vtkGitConvertWinPath = Chr(34) & convertedPath & Chr(34)
    Exit Function
    

vtkGitConvertWinPath_Error:
    If (Err.Number = VTK_FORBIDDEN_PARAMETER) Then
        Err.Raise Err.Number, "Module vtkGitFunctions ; Function vtkGitConvertWinPath", Err.Description
    Else
        'Debug.Print "ERR IN CONVERT : " & Err.Number & Err.Description
        Err.Raise VTK_UNEXPECTED_ERROR, "Module vtkGitFunctions ; Function vtkGitConvertWinPath", Err.Description
    End If
    Exit Function

End Function
