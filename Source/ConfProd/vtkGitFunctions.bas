Attribute VB_Name = "vtkGitFunctions"
'---------------------------------------------------------------------------------------
' Module    : vtkGitHubFunctions
' Author    : Abdelfattah Lahbib
' Date      : 08/06/2013
' Purpose   : Interact with Git
'---------------------------------------------------------------------------------------

Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : isGitCmdInString
' Author    : Lucas Vitorino
' Purpose   : Test if a string contains the "Git\cmd" substring.
'---------------------------------------------------------------------------------------
'
Private Function isGitCmdInString(myString As String) As Boolean
    isGitCmdInString = (InStr(UCase(myString), UCase("Git\cmd")))
End Function


'---------------------------------------------------------------------------------------
' Procedure : isGitInstalled
' Author    : Abdelfattah Lahbib
' Date      : 09/06/2013
' Purpose   : - Check if the Git "cmd.exe" is accessible via the PATH.
'---------------------------------------------------------------------------------------
'
Private Function isGitInstalled() As Boolean
    
    'Test if the "Git\cmd" substring is in the PATH string
    isGitInstalled = InStr(UCase(Environ("PATH")), UCase("Git\cmd"))
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeGit
' Author    : Abdelfattah Lahbib
' Date      : 27/04/2013
' Purpose   : - Initialize Git in a directory using the "git init" command.
'             - Optionally, writes the output of the "git init" command in a textfile,
'               in a subfolder whose name has to be passed as parameter.
'---------------------------------------------------------------------------------------
'
Public Function vtkInitializeGit(folderPath As String, Optional textLogFolderName As String = "", _
                                                       Optional textLogFileName As String = "initialize.log")
    
    Debug.Print "0"
    Dim FileInitPath As String
    Dim retShell As String
 
    On Error GoTo vtkInitializeGit_Err
     
    If isGitInstalled = False Then
        err.Raise 3000, "", "Git is not installed"
    End If
    
    If vtkDoesFolderExist(folderPath & "\.git") Then
        err.Raise 3100, "", "Git has already been initialized in the folder " & folderPath
    End If
    
    If textLogFolderName = "" Then
        ' No text log folder specified, simple git init
        retShell = Shell("cmd.exe /c git init " & vtkGitConvertWinPath(folderPath))
        Application.Wait (Now + TimeValue("0:00:01"))
    Else
        ' Text log folder specified
        ' If it does not yet exist, create text log folder
        If vtkDoesFolderExist(folderPath & "\" & textLogFolderName) = False Then
            MkDir folderPath & "\" & textLogFolderName
        End If
        ' If it does not yet exist, create text log file
        If Dir(folderPath & "\" & textLogFolderName & "\" & textLogFileName) = "" Then
            Dim fso As New FileSystemObject
            fso.CreateTextFile (folderPath & "\" & textLogFolderName & "\" & textLogFileName)
        End If
        ' Git init with redirection of the output
        retShell = Shell("cmd.exe /c git init " & vtkGitConvertWinPath(folderPath) & "  > " _
        & vtkGitConvertWinPath(folderPath & "\" & textLogFolderName & "\" & textLogFileName))
        Application.Wait (Now + TimeValue("0:00:01"))
    End If
    
    On Error GoTo 0
    vtkInitializeGit = 0
    Exit Function
    
    End If
    
    
vtkInitializeGit_Err:
    Debug.Print "Error " & err.Number & " in vtkInitializeGit : " & err.Description
    vtkInitializeGit = err.Number
    Exit Function

End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkStatusGit
' Author    : Abdelfattah Lahbib
' Date      : 30/04/2013
' Purpose   : - Log the output of the "git status" command in $projectDirectory\GitLog\logStatus.log
'             - Pop a MsgBox with the content of this file.
'---------------------------------------------------------------------------------------
'
Public Function vtkStatusGit() As String
  
    Dim ActiveProjPath As String
    Dim logFileName As String
    Dim GitLogRelativeDirPath As String
    Dim GitLogAbsoluteDirPath As String
    Dim logFileFullPath As String
    Dim retShell As String
    Dim retShell2 As String
    Dim fso As New FileSystemObject
      
    On Error GoTo vtkStatusGit_Error
    
    ' make paths
    ActiveProjPath = fso.GetParentFolderName(ActiveWorkbook.path)
    logFileName = "logStatus.log"
    GitLogRelativeDirPath = "GitLog"
    GitLogAbsoluteDirPath = ActiveProjPath & "\" & GitLogRelativeDirPath
    logFileFullPath = GitLogAbsoluteDirPath & "\" & logFileName
    
    ' create status log file
    vtkCreateFileInDirectory logFileName, GitLogAbsoluteDirPath
    ' Execute git status in the project directory  and log the output in the relevant file
    retShell = Shell("cmd.exe /k cd " & ActiveProjPath & " & git status   >" & logFileFullPath & " ", vbHide)
    ' make a break to execute shell commands
    Application.Wait (Now + TimeValue("0:00:01"))
    ' kill related processus
    retShell2 = Shell("cmd.exe /k  TASKKILL /IM cmd.exe")
    ' make a break to execute shell commands
    Application.Wait (Now + TimeValue("0:00:01"))
    ' read log file , and return its content
    vtkStatusGit = vtkTextFileReader(logFileFullPath)

   On Error GoTo 0
   Exit Function

vtkStatusGit_Error:
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure vtkStatusGit of Module vtkGitFunctions"
    vtkStatusGit = err.Number
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkGitConvertWinPath
' Author    : Lucas Vitorino
' Purpose   : Converts an *absolute* Windows PATH to a one suitable for use with Git.
' Notes     : - WIP because I don't fully understand the behaviour of Git. I suppose
'               it's a Windows/Unix path format conflict.
'             - For now, I can't specify the drive letter in a path used with Git.
'               This function will strip it and git will assume it's "C" .
'---------------------------------------------------------------------------------------
'
Public Function vtkGitConvertWinPath(winPath As String) As String
    
    Dim unixPath As String
    'Changing the backslahes in slashes
    ' NB : Optional
    unixPath = Replace(winPath, "\", "/")
    
    'Removing the drive letter and the semicolon in the beginning
    unixPath = Replace(unixPath, Left(unixPath, 1), "", 1, 1)
    unixPath = Replace(unixPath, Left(unixPath, 1), "", 1, 1)
    
    vtkGitConvertWinPath = Chr(34) & unixPath & Chr(34)
    
End Function
