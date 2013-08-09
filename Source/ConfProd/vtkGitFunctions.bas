Attribute VB_Name = "vtkGitFunctions"
'---------------------------------------------------------------------------------------
' Module    : vtkGitHubFunctions
' Author    : Abdelfattah Lahbib
' Date      : 08/06/2013
' Purpose   : Interact with Git
'---------------------------------------------------------------------------------------

Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : vtkIsGitCmdInString
' Author    : Lucas Vitorino
' Purpose   : Test if a string contains the "Git\cmd" substring.
'---------------------------------------------------------------------------------------
'
Public Function vtkIsGitCmdInString(myString As String)
    
On Error GoTo vtkIsGitCmdInString_Error

    vtkIsGitCmdInString = False
    If (InStr(UCase(myString), UCase("Git\cmd"))) Then vtkIsGitCmdInString = True

   On Error GoTo 0
   Exit Function

vtkIsGitCmdInString_Error:
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure vtkIsGitCmdInString of Module vtkGitFunctions"

End Function




'---------------------------------------------------------------------------------------
' Procedure : vtkVerifyEnvirGitVar
' Author    : Abdelfattah Lahbib
' Date      : 09/06/2013
' Purpose   : - Check if the Git "cmd.exe" is accessible via the PATH.
'             - If not, pop a MsgBox.
'---------------------------------------------------------------------------------------
'
Public Function vtkVerifyEnvirGitVar() As String
    
    Dim retVal As String
    
On Error GoTo vtkVerifyEnvirGitVar_Error
      
    'Test if the "Git\cmd" substring is in the PATH string
    If vtkIsGitCmdInString(Environ("PATH")) Then
        vtkVerifyEnvirGitVar = ""
    Else
        vtkVerifyEnvirGitVar = "problem"
        MsgBox "Error : Git is not accessible via your path." & vbCrLf & vbCrLf & _
        "To correct the problem, see the tutorial on : https://github.com/jpimbert/VBAToolKit/wiki/VbaToolKit-SetUp " & vbCrLf & vbCrLf & _
        "You local repository has not been initialized", vbInformation
    End If

    On Error GoTo 0
    Exit Function

vtkVerifyEnvirGitVar_Error:
    Debug.Print "Error " & err.Number & " in vtkVerifyEnvirGitVar : " & err.Description
    vtkVerifyEnvirGitVar = err.Number
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeGit
' Author    : Abdelfattah Lahbib
' Date      : 27/04/2013
' Purpose   : - Initialize Git in the project root folder using the "git init" command.
'             - Log the output of the "git init" command in $projectDirectory\GitLog\logGitInitialize.log
'---------------------------------------------------------------------------------------
'
Public Function vtkInitializeGit()
       
    Dim logFileName As String
    Dim GitLogAbsoluteDirPath As String
    Dim GitLogRelativeDirPath As String
    Dim FileInitPath As String
    Dim RetShell1 As String
    Dim RetShell2 As String
    Dim RetFnVerifyEnvVar As String
    Dim ActiveProjPath As String
    Dim RetShellMessage As String
    Dim logFileFullPath As String
    Dim fso As New FileSystemObject
 
    On Error GoTo vtkInitializeGit_Err
 
    RetFnVerifyEnvVar = vtkVerifyEnvirGitVar()
     
 If RetFnVerifyEnvVar <> "problem" Then
 
    ' Make paths
    ActiveProjPath = fso.GetParentFolderName(ActiveWorkbook.path)
    GitLogRelativeDirPath = "GitLog"
    GitLogAbsoluteDirPath = ActiveProjPath & "\" & GitLogRelativeDirPath
    logFileName = "logGitInitialize.log"
  
    ' Create log file
    logFileFullPath = vtkCreateFileInDirectory(logFileName, GitLogAbsoluteDirPath)
    ' Execute git init in the project directory and log the output in the relevant file
    ' NB : You have to redirect the log using the *relative* path.
    RetShell1 = Shell("cmd.exe /k cd " & ActiveProjPath & " & git init   >" & GitLogRelativeDirPath & "\" & logFileName & " ")
    ' Make a break to execute shell commands
    Application.Wait (Now + TimeValue("0:00:01"))
    ' Kill related processus
    RetShell2 = Shell("cmd.exe /k  TASKKILL /IM cmd.exe")
    ' Make a break to execute shell commands
    Application.Wait (Now + TimeValue("0:00:01"))

 End If
 
 On Error GoTo 0
 Exit Function
 
vtkInitializeGit_Err:
    Debug.Print "Error " & err.Number & " in vtkInitializeGit : " & err.Description
    vtkInitializeGit = err.Number
 
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
    Dim RetShell As String
    Dim RetShell2 As String
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
    RetShell = Shell("cmd.exe /k cd " & ActiveProjPath & " & git status   >" & logFileFullPath & " ", vbHide)
    ' make a break to execute shell commands
    Application.Wait (Now + TimeValue("0:00:01"))
    ' kill related processus
    RetShell2 = Shell("cmd.exe /k  TASKKILL /IM cmd.exe")
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
' Procedure : vtkCreateFileInDirectory
' Author    : Abdelfattah Lahbib
' Date      : 30/04/2013
' Purpose   : Create a file named $fileName in the directory $directory
' Notes     : Notably used for creating Git log files
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateFileInDirectory(fileName As String, directory As String) As String

    Dim fso As New FileSystemObject
    Dim fullFilePath As String
      
On Error GoTo vtkCreateFileInDirectory_Error
    
    fullFilePath = directory & "\" & fileName
      
    ' If the file doesn't exist, we create it
    If fso.FileExists(fullFilePath) = False Then
            fso.CreateTextFile (fullFilePath)
    End If
      
    'return full created file path
    vtkCreateFileInDirectory = fullFilePath

    On Error GoTo 0
    Exit Function

vtkCreateFileInDirectory_Error:
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure vtkCreateFileInDirectory of Module vtkGitFunctions"
    vtkCreateFileInDirectory = err.Number

End Function

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
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure VtkTextFileReader of Module vtkGitFunctions"
    vtkTextFileReader = err.Number
    
End Function



