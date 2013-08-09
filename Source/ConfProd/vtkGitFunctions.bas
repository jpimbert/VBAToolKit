Attribute VB_Name = "vtkGitFunctions"
'---------------------------------------------------------------------------------------
' Module    : vtkGitHubFunctions
' Author    : Abdelfattah Lahbib
' Date      : 08/06/2013
' Purpose   : Interact with Git
'---------------------------------------------------------------------------------------

Option Explicit



'---------------------------------------------------------------------------------------
' Procedure : vtkVerifyEnvirGitVar
' Author    : Abdelfattah Lahbib
' Date      : 09/06/2013
' Purpose   : Check if the Git "cmd.exe" is accessible via the PATH.
'---------------------------------------------------------------------------------------
'
Public Function vtkVerifyEnvirGitVar() As String
    
    On Error GoTo vtkVerifyEnvirGitVar_Err
    
    Dim EnvString As String
    Dim retval As String
      
    EnvString = Environ("PATH")
    'Test if the "Git\cmd" substring is in the PATH string
    If (InStr(UCase(EnvString), UCase("Git\cmd"))) Then
        vtkVerifyEnvirGitVar = ""
    Else
        vtkVerifyEnvirGitVar = "problem"
        MsgBox "Error : Git is not accessible via your path." & vbCrLf & vbCrLf & _
        "To correct the problem, see the tutorial on : https://github.com/jpimbert/VBAToolKit/wiki/VbaToolKit-SetUp " & vbCrLf & vbCrLf & _
        "You local repository has not been initialized", vbInformation
    End If

    On Error GoTo 0
    Exit Function

vtkVerifyEnvirGitVar_Err:
    Debug.Print "Error " & err.Number & " in vtkVerifyEnvirGitVar : " & err.Description
    vtkVerifyEnvirGitVar = err.Number
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeGit
' Author    : Abdelfattah Lahbib
' Date      : 27/04/2013
' Purpose   :- create file to contain command result
'            - verify git path
'            - return git path
'            - HKEY_CLASSES_ROOT\github-windows\shell\open\command
'---------------------------------------------------------------------------------------
'
Public Function vtkInitializeGit()
       
    Dim GitLogFileName As String
    Dim GitLogAbsoluteDirPath As String
    Dim GitLogRelativeDirPath As String
    Dim FileInitPath As String
    Dim RetShell1 As String
    Dim RetShell2 As String
    Dim RetFnVerifyEnvVar As String
    Dim ActiveProjPath As String
    Dim RetShellMessage As String
    Dim FullFileLogPath As String
    Dim fso As New FileSystemObject
 
    On Error GoTo vtkInitializeGit_Err
 
    RetFnVerifyEnvVar = vtkVerifyEnvirGitVar()
 
 If RetFnVerifyEnvVar <> "problem" Then
    ' Make paths
    ActiveProjPath = fso.GetParentFolderName(ActiveWorkbook.path)
    GitLogRelativeDirPath = "GitLog"
    GitLogAbsoluteDirPath = ActiveProjPath & "\" & GitLogRelativeDirPath
    GitLogFileName = "logGitInitialize.log"
  
    ' Create log file
    FullFileLogPath = vtkcreatefilegit(GitLogFileName, GitLogAbsoluteDirPath)
    ' Execute shell commands
    ' NB : You have to redirect the log using the *relative* path.
    RetShell1 = Shell("cmd.exe /k cd " & ActiveProjPath & " & git init   >" & GitLogRelativeDirPath & "\" & GitLogFileName & " ")
    ' Make a break to execute shell commands
    Application.Wait (Now + TimeValue("0:00:01"))
    ' Kill related processus
    RetShell2 = Shell("cmd.exe /k  TASKKILL /IM cmd.exe")
    ' Make a break to execute shell commands
    Application.Wait (Now + TimeValue("0:00:01"))
    ' Read log file
    RetShellMessage = VtkFileReader(GitLogFileName, GitLogAbsoluteDirPath)

 End If
 
 On Error GoTo 0
 vtkInitializeGit = 0
 Exit Function
 
vtkInitializeGit_Err:
    Debug.Print "Error " & err.Number & " in vtkInitializeGit : " & err.Description
    vtkInitializeGit = err.Number
 
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkStatusGit
' Author    : user
' Date      : 30/04/2013
' Purpose   : -this function execute shell command to take git status
'             -write a text file contain the command result
'             -return message on the functiuon name
'
'---------------------------------------------------------------------------------------
'
Public Function vtkStatusGit() As String
  
    Dim GitStatusFileName As String
    
    Dim PathOfGitStatusFile As String
    Dim ActiveProjPath As String
    Dim GitLogFilePath As String
    Dim GitDir As String
    Dim RetShell As String
    Dim RetShell2 As String
    Dim fso As New FileSystemObject
      
    ' make paths
    ActiveProjPath = fso.GetParentFolderName(ActiveWorkbook.path)
    GitStatusFileName = "\logStatus.log"
    GitDir = ActiveProjPath & "\GitLog"
    ' create status log file
    PathOfGitStatusFile = vtkcreatefilegit(GitStatusFileName, GitDir)
    ' execute shell command
    RetShell = Shell("cmd.exe /k cd " & ActiveProjPath & " & git status   >" & GitDir & GitStatusFileName & " ", vbHide)
    ' make a break to execute shell commands
    Application.Wait (Now + TimeValue("0:00:01"))
    ' kill related processus
    RetShell2 = Shell("cmd.exe /k  TASKKILL /IM cmd.exe")
    ' make a break to execute shell commands
    Application.Wait (Now + TimeValue("0:00:01"))
    ' read log file , and return it in function name
    vtkStatusGit = VtkFileReader(GitStatusFileName, GitDir)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkcreatefilegit
' Author    : user
' Date      : 30/04/2013
' Purpose   : - this function will create files that will be contain console command result
'             - tested
'---------------------------------------------------------------------------------------
'
Public Function vtkcreatefilegit(FileName As String, GitFolderPath As String) As String

 Dim fso As New FileSystemObject
 Dim FullFilePath As String
  'make full file path
  FullFilePath = GitFolderPath & FileName
  'if log file don't exist we will create it
  If fso.FileExists(FullFilePath) = False Then
        fso.CreateTextFile (FullFilePath)
  End If
  'return full created file path
  vtkcreatefilegit = FullFilePath
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkfilereader
' Author    : Abdelfattah Lahbib
' Date      : 30/04/2013
' Purpose   : - take file name and projectpath on parameters
'             - return the text on the file on function name
'---------------------------------------------------------------------------------------
'
Public Function VtkFileReader(FileName As String, ProjectGitPath As String) As String

    Dim Textfile As Variant
    Dim strresult As String
    Dim fso As New FileSystemObject
    Dim FullFilePath As String

    FullFilePath = (ProjectGitPath & FileName)

    Set Textfile = fso.OpenTextFile(FullFilePath, ForReading)
    'while not end of file
    Do Until Textfile.AtEndOfStream
    'read line per line
        strresult = strresult & Chr(10) & Textfile.ReadLine
    Loop
    'return file text
    VtkFileReader = strresult

End Function



