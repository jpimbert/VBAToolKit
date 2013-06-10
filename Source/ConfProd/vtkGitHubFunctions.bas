Attribute VB_Name = "vtkGitHubFunctions"
'---------------------------------------------------------------------------------------
' Module    : vtkGitHubFunctions
' Author    : user
' Date      : 08/06/2013
' Purpose   :- verifier que la variable d'environnement est deja ajouter
'            - initialiser le repesitory git
'---------------------------------------------------------------------------------------

Option Explicit
Public Function vtkGitFolderSetter() As String
    
    Dim retval As Variant
    Dim fldname As String
    Dim fd As FileDialog, fl As Variant
    'path picker
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        With fd
            .AllowMultiSelect = False
            .Title = "Please select ..\Git\cmd folder"
            .Show
        End With
    
        For Each fl In fd.SelectedItems
        fldname = fl
        Next
     'if user will choose a bad or an empty path function will return null
    vtkGitFolderSetter = ""
    
    If InStr(UCase(fldname), UCase("Git\cmd")) Then
        
        vtkGitFolderSetter = fldname
        MsgBox "ok", vbInformation
        
    ElseIf fldname = "" Then
        
        MsgBox "Folder Not selected ,try to set envirenement var manually ,  ", vbInformation, "Note:"
    
    Else
        MsgBox "wrong folder is selected ", vbInformation, "Note:"
        retval = vtkGitFolderSetter()
   End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkVerifyEnvirGitVar
' Author    : user
' Date      : 09/06/2013
' Purpose   : - verify if environement variable is active ,
'             - allow user to try to define envirenement var
'---------------------------------------------------------------------------------------
'
Public Function vtkVerifyEnvirGitVar() As String

Dim EnvString As String
Dim retval As String
  
  EnvString = Environ("PATH")
  
   'test if git environement var already exist
   If (InStr(UCase(EnvString), UCase("Git\cmd"))) Then
       'var already defined  return ""
       vtkVerifyEnvirGitVar = ""
   Else
    ' var don't exist , allow to user to try to define it
    retval = vtkGitFolderSetter()
       'treat user choice :
        If retval <> "" Then
           'user was define the good location
           vtkVerifyEnvirGitVar = retval
        Else
           'user was define a false location
           vtkVerifyEnvirGitVar = "problem"
        End If
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeGit
' Author    : user
' Date      : 27/04/2013
' Purpose   :- create file to contain command result
'            - verify git path
'            - return git path
'            - HKEY_CLASSES_ROOT\github-windows\shell\open\command
'---------------------------------------------------------------------------------------
'
Public Function vtkInitializeGit(project As String) As String
    Dim proje As vtkProject
    Set proje = vtkProjectForName(projectName:=project)
    
    Dim InitFileName As String
    
    Dim vtkGitFiles As String
    Dim FileInitPath As String
    Dim RetVal1 As Variant
    Dim retval As String
    Dim pathactiveproj As String
 
 PathActiveProj = proje.
 
 'function how will verify envirenemnt var
    retval = vtkVerifyEnvirGitVar()
 
 If retval <> "problem" Then
 
  vtkGitFiles = fso.GetParentFolderName(ActiveWorkbook.path) & "\GitLog"
  InitFileName = "\logGitInitialize.log"
  'create log file
   FileInitPath = vtkcreatefilegit(InitFileName, vtkGitFiles)

  RetVal1 = Shell("cmd.exe /k cd " & pathactiveproj & " && path =" & retval & ";%path% & git init   >" & FileInitPath & " ", vbNormalFocus)
 
 
 vtkInitializeGit = VtkFileReader(InitFileName, vtkGitFiles)
 
 End If
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
  
  Dim StatusFileName As String
  StatusFileName = "\logStatus.txt"

 PathOfFileStatus = vtkcreatefilegit(StatusFileName, vtkprojectpathtestgit)

 RetVal2 = Shell("cmd.exe /k cd " & vtkprojectpathtestgit & " & git status   >" & PathOfFileStatus & " ", vbHide)

'Debug.Print strresult
 vtkStatusGit = VtkFileReader(StatusFileName, vtkprojectpathtestgit)
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkcreatefilegit
' Author    : user
' Date      : 30/04/2013
' Purpose   : - this function will create files that will be contain console command result
'             - tested
'---------------------------------------------------------------------------------------
'
Public Function vtkcreatefilegit(FileName As String, ProjectGitPath As String) As String

 Dim fso As New FileSystemObject
 Dim FullFilePath As String

  FullFilePath = ProjectGitPath & FileName

  If fso.FileExists(FullFilePath) = False Then
        fso.CreateTextFile (FullFilePath)
  End If
    
    
  vtkcreatefilegit = FullFilePath
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkfilereader
' Author    : user
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
Do Until Textfile.AtEndOfStream
    strresult = strresult & Textfile.ReadLine
Loop
VtkFileReader = strresult
'Debug.Print strresult
End Function

