Attribute VB_Name = "vtkGitFunctions"
'---------------------------------------------------------------------------------------
' Module    : vtkGitHubFunctions
' Author    : Abdelfattah Lahbib
' Date      : 08/06/2013
' Purpose   : Interact with Git
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeGit
' Author    : Abdelfattah Lahbib
' Date      : 27/04/2013
' Purpose   : - Initialize Git in a directory using the "git init" command.
'               The directory path has to be absolute, and on the C: drive
'             - Optionally, writes the output of the "git init" command in a textfile
'               whose path has to be passed as parameter.
'               If the path is absolute, a drive other than C: will raise an error
'               If the path is relative, it will be considered as relatie to the folder path.
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
    Dim FileInitPath As String
    Dim retShell As String
    Dim tmpLogFileName As String
    tmpLogFileName = "initialize.log"
    Dim logFileFullPath As String
    
    Dim convertedFolderPath As String
    Dim convertedLogFilePath As String
    
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
    retShell = ShellAndWait("cmd.exe /c git init " & convertedFolderPath _
    & " > " & convertedLogFilePath, 0, vbHide, AbandonWait)
    
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
        'Debug.Print "ERR IN INITIALIZE : " & err.Number & err.Description
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
